import asyncio
from collections.abc import Callable
from typing import Generic, TypeVar

from agent_server.tools.list_files import VOLUME_URI
from agent_server.utils import get_user_workspace_client

T = TypeVar("T")


class VolumeFileCache(Generic[T]):
    """Downloads files from the volume, parses them, and keeps the parsed result in memory.

    The files in the volume do not change, so a file stays cached until the app restarts
    or until it is evicted to make room for another file.
    """

    def __init__(self, extensions: tuple[str, ...], parse: Callable[[bytes], T], max_entries: int):
        self.extensions = extensions
        self.parse = parse
        self.max_entries = max_entries
        self._entries: dict[str, T] = {}
        self._lock = asyncio.Lock()

    async def get(self, filename: str) -> T:
        if "/" in filename:
            raise ValueError("Pass only the file name, without a path.")
        if not filename.lower().endswith(self.extensions):
            raise ValueError(f"Unsupported file type: {filename}. Supported types: {', '.join(self.extensions)}.")

        # The lock makes parallel tool calls for the same file wait for one download
        async with self._lock:
            if filename in self._entries:
                return self._entries[filename]

            client = get_user_workspace_client()
            response = await asyncio.to_thread(client.files.download, f"{VOLUME_URI}/{filename}")
            if response.contents is None:
                raise ValueError(f"The file {filename} is empty.")
            data = await asyncio.to_thread(response.contents.read)
            parsed = await asyncio.to_thread(self.parse, data)

            while len(self._entries) >= self.max_entries:
                self._entries.pop(next(iter(self._entries)))
            self._entries[filename] = parsed
            return parsed
