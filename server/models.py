"""Fractionate Edge — Model Managers

LlamaServerManager: manages the llama-server subprocess for Falcon3-7B text inference.
Florence2Manager: manages in-process loading/unloading of Florence-2-base for vision tasks.
"""

import asyncio
import gc
import logging
import os
import subprocess
import time
from pathlib import Path
from typing import Optional

import httpx

logger = logging.getLogger("fractionate.models")


class LlamaServerManager:
    """Manages the llama-server subprocess for BitNet text inference."""

    def __init__(self, binary_path: str, model_path: str, port: int = 8082):
        self.binary_path = binary_path
        self.model_path = model_path
        self.port = port
        self.process: Optional[subprocess.Popen] = None
        self._start_time: Optional[float] = None

    async def start(self, threads: Optional[int] = None, context_size: int = 4096):
        """Launch llama-server as a subprocess."""
        if self.is_running:
            logger.info("llama-server is already running (pid=%s)", self.process.pid)
            return

        if not os.path.isfile(self.binary_path):
            raise FileNotFoundError(f"llama-server binary not found: {self.binary_path}")
        if not os.path.isfile(self.model_path):
            raise FileNotFoundError(f"Model file not found: {self.model_path}")

        thread_count = threads or self._optimal_threads()
        cmd = [
            self.binary_path,
            "-m", self.model_path,
            "--host", "127.0.0.1",
            "--port", str(self.port),
            "-t", str(thread_count),
            "-c", str(context_size),
            "--log-disable",
        ]

        logger.info("Starting llama-server: %s", " ".join(cmd))
        self.process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        self._start_time = time.time()

        try:
            await self._wait_for_ready()
            logger.info("llama-server started successfully (pid=%s)", self.process.pid)
        except TimeoutError:
            await self.stop()
            raise

    async def stop(self):
        """Kill the llama-server subprocess."""
        if self.process is None:
            return

        pid = self.process.pid
        logger.info("Stopping llama-server (pid=%s)", pid)

        self.process.terminate()
        try:
            self.process.wait(timeout=10)
        except subprocess.TimeoutExpired:
            logger.warning("llama-server did not terminate gracefully, killing (pid=%s)", pid)
            self.process.kill()
            self.process.wait(timeout=5)

        self.process = None
        self._start_time = None
        logger.info("llama-server stopped")

    @property
    def is_running(self) -> bool:
        return self.process is not None and self.process.poll() is None

    @property
    def pid(self) -> Optional[int]:
        if self.is_running:
            return self.process.pid
        return None

    @property
    def uptime_seconds(self) -> Optional[float]:
        if self.is_running and self._start_time:
            return time.time() - self._start_time
        return None

    async def health_check(self) -> bool:
        """Check if llama-server is responding to health requests."""
        if not self.is_running:
            return False
        try:
            async with httpx.AsyncClient(timeout=5.0) as client:
                r = await client.get(f"http://127.0.0.1:{self.port}/health")
                return r.status_code == 200
        except (httpx.ConnectError, httpx.TimeoutException):
            return False

    def _optimal_threads(self) -> int:
        """Use physical cores (not logical/hyperthreaded) for best throughput."""
        return max(1, (os.cpu_count() or 4) // 2)

    async def _wait_for_ready(self, timeout: int = 120):
        """Poll health endpoint until llama-server responds."""
        async with httpx.AsyncClient() as client:
            deadline = time.time() + timeout
            while time.time() < deadline:
                # Check if process died
                if self.process.poll() is not None:
                    stderr = self.process.stderr.read().decode(errors="replace") if self.process.stderr else ""
                    raise RuntimeError(
                        f"llama-server exited with code {self.process.returncode}: {stderr[:500]}"
                    )
                try:
                    r = await client.get(f"http://127.0.0.1:{self.port}/health")
                    if r.status_code == 200:
                        return
                except httpx.ConnectError:
                    pass
                await asyncio.sleep(0.5)
            raise TimeoutError(f"llama-server failed to start within {timeout}s")

    def get_status(self) -> dict:
        """Return status dict for the health endpoint."""
        return {
            "installed": os.path.isfile(self.binary_path) and os.path.isfile(self.model_path),
            "running": self.is_running,
            "model_path": self.model_path,
            "pid": self.pid,
            "uptime_seconds": self.uptime_seconds,
        }


class Florence2Manager:
    """Manages in-process loading/unloading of Florence-2-base for vision tasks."""

    def __init__(self, model_path: str):
        self.model_path = model_path
        self.model = None
        self.processor = None
        self.is_loaded = False
        self._load_time: Optional[float] = None

    def load(self):
        """Load Florence-2 model and processor into RAM."""
        if self.is_loaded:
            logger.info("Florence-2 is already loaded")
            return

        logger.info("Loading Florence-2 from %s", self.model_path)
        start = time.time()

        from transformers import AutoModelForCausalLM, AutoProcessor

        self.model = AutoModelForCausalLM.from_pretrained(
            self.model_path, trust_remote_code=True
        )
        self.processor = AutoProcessor.from_pretrained(
            self.model_path, trust_remote_code=True
        )
        self.model.eval()
        self.is_loaded = True
        self._load_time = time.time()

        elapsed = time.time() - start
        logger.info("Florence-2 loaded in %.1fs", elapsed)

    def unload(self):
        """Free RAM by unloading the model."""
        if not self.is_loaded:
            return

        logger.info("Unloading Florence-2")
        del self.model
        del self.processor
        self.model = None
        self.processor = None
        self.is_loaded = False
        self._load_time = None

        gc.collect()

        try:
            import torch
            if torch.cuda.is_available():
                torch.cuda.empty_cache()
        except ImportError:
            pass

        logger.info("Florence-2 unloaded")

    def process_image(self, image) -> dict:
        """Process a single PIL Image and return structured text with caption and OCR.

        Args:
            image: PIL.Image.Image instance

        Returns:
            dict with 'caption' and 'ocr' keys containing extracted information
        """
        if not self.is_loaded:
            self.load()

        import torch

        results = {}

        # Detailed caption — understands layout, images, structure
        prompt = "<MORE_DETAILED_CAPTION>"
        inputs = self.processor(text=prompt, images=image, return_tensors="pt")
        with torch.no_grad():
            generated = self.model.generate(**inputs, max_new_tokens=1024)
        caption_text = self.processor.batch_decode(generated, skip_special_tokens=False)[0]
        results["caption"] = self.processor.post_process_generation(
            caption_text, task=prompt, image_size=image.size
        )

        # OCR with region detection — text + bounding boxes
        prompt = "<OCR_WITH_REGION>"
        inputs = self.processor(text=prompt, images=image, return_tensors="pt")
        with torch.no_grad():
            generated = self.model.generate(**inputs, max_new_tokens=1024)
        ocr_text = self.processor.batch_decode(generated, skip_special_tokens=False)[0]
        results["ocr"] = self.processor.post_process_generation(
            ocr_text, task=prompt, image_size=image.size
        )

        return results

    def get_status(self) -> dict:
        """Return status dict for the health endpoint."""
        installed = os.path.isdir(self.model_path)
        return {
            "installed": installed,
            "loaded": self.is_loaded,
            "model_path": self.model_path,
        }
