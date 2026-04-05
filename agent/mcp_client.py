"""
MCP Client – Robust connection handling for Windows
"""

import asyncio
import json
import os
import re
import sys
from typing import Any, Optional

_MCP_IMPORT_ERROR: str | None = None
ClientSession = None
StdioServerParameters = None
stdio_client = None

if sys.version_info >= (3, 14):
    _MCP_IMPORT_ERROR = (
        "MCP is not stable on Python 3.14 in this environment. "
        "Please use Python 3.11 or 3.12 for this project."
    )
else:
    try:
        from mcp import ClientSession, StdioServerParameters
        from mcp.client.stdio import stdio_client
    except Exception as e:
        _MCP_IMPORT_ERROR = f"Failed to import MCP dependencies: {e}"


def _extract_json(text: str):
    """Extracts JSON from markdown or raw text."""
    if not isinstance(text, str):
        return None
    text = re.sub(r"```[a-z]*\n?", "", text).replace("```", "").strip()
    try:
        return json.loads(text)
    except Exception:
        pass
    for opener, closer in [("[", "]"), ("{", "}")]:
        s, e = text.find(opener), text.rfind(closer)
        if s != -1 and e > s:
            try:
                return json.loads(text[s: e + 1])
            except Exception:
                pass
    return None


# Path resolution
_THIS_DIR = os.path.dirname(os.path.realpath(__file__))
_PROJECT_ROOT = os.path.dirname(_THIS_DIR)
SERVER_SCRIPT = os.path.join(_PROJECT_ROOT, "mcp_servers", "ppt_server.py")


class MCPClient:
    """Robust MCP Client with proper cleanup."""
    
    def __init__(self):
        self._session: Optional[Any] = None
        self._read = None
        self._write = None
        self._stdio_context = None
        self._session_context = None
        self._local_backend: Optional[Any] = None
        self._mode = "mcp"

    def _init_local_backend(self) -> None:
        from agent.local_ppt_backend import LocalPPTBackend

        self._local_backend = LocalPPTBackend()
        
    async def __aenter__(self):
        """Initialize connection."""
        if _MCP_IMPORT_ERROR is not None:
            self._mode = "local"
            self._init_local_backend()
            return self

        # Verify server exists
        if not os.path.exists(SERVER_SCRIPT):
            raise FileNotFoundError(f"MCP server not found: {SERVER_SCRIPT}")
        
        # Create server parameters
        server_params = StdioServerParameters(
            command=sys.executable,
            args=[SERVER_SCRIPT],
            env={**os.environ},
        )
        
        try:
            # Create stdio client context
            self._stdio_context = stdio_client(server_params)
            self._read, self._write = await self._stdio_context.__aenter__()
            
            # Create session context
            self._session_context = ClientSession(self._read, self._write)
            self._session = await self._session_context.__aenter__()
            
            # Initialize session
            await self._session.initialize()
            
            return self
            
        except Exception as e:
            await self.__aexit__(type(e), e, e.__traceback__)
            self._mode = "local"
            self._init_local_backend()
            return self
    
    async def __aexit__(self, exc_type, exc_val, exc_tb):
        """Clean up with proper error handling."""
        if self._mode == "local":
            self._local_backend = None
            return

        errors = []
        
        # Close session first
        if self._session_context is not None:
            try:
                await self._session_context.__aexit__(exc_type, exc_val, exc_tb)
            except Exception as e:
                errors.append(f"Session error: {e}")
            self._session_context = None
            self._session = None
        
        # Close stdio connection
        if self._stdio_context is not None:
            try:
                await self._stdio_context.__aexit__(exc_type, exc_val, exc_tb)
            except Exception as e:
                errors.append(f"Stdio error: {e}")
            self._stdio_context = None
            self._read = None
            self._write = None
        
        # Don't propagate cleanup errors if we're already handling an exception
        if errors and exc_type is None:
            print(f"Cleanup warnings: {errors}")
    
    async def call_tool(self, name: str, arguments: dict[str, Any] | None = None) -> str:
        """Call an MCP tool."""
        if self._mode == "local":
            if self._local_backend is None:
                raise RuntimeError("Local backend not initialized")

            args = arguments or {}
            if name == "create_presentation":
                return self._local_backend.create_presentation()
            if name == "add_slide":
                return self._local_backend.add_slide(args.get("title", "Untitled"), args.get("bullets", []))
            if name == "add_image_slide":
                return self._local_backend.add_image_slide(
                    title=args.get("title", "Image Slide"),
                    image_prompt=args.get("image_prompt", ""),
                    caption=args.get("caption", ""),
                )
            if name == "save_presentation":
                return self._local_backend.save_presentation(args.get("filename", "output.pptx"))
            if name == "plan_slides":
                return self._local_backend.plan_slides(
                    topic=args.get("topic", "Presentation Topic"),
                    num_slides=args.get("num_slides", 5),
                )
            if name == "generate_content":
                return self._local_backend.generate_content(args.get("title", "Slide"))

            raise RuntimeError(f"Unknown tool: {name}")

        if self._session is None:
            raise RuntimeError("Not connected. Use 'async with MCPClient() as client:'")
        
        try:
            result = await self._session.call_tool(name, arguments or {})
            return "\n".join(block.text for block in result.content if hasattr(block, "text"))
        except Exception as e:
            raise RuntimeError(f"Tool '{name}' failed: {e}")
    
    async def list_tools(self) -> list[str]:
        """List available tools."""
        if self._mode == "local":
            if self._local_backend is None:
                raise RuntimeError("Local backend not initialized")
            return self._local_backend.list_tools()

        if self._session is None:
            raise RuntimeError("Not connected")
        
        result = await self._session.list_tools()
        return [t.name for t in result.tools]
    
    async def create_presentation(self) -> str:
        """Create a new presentation."""
        return await self.call_tool("create_presentation")
    
    async def add_slide(self, title: str, bullets: list[str]) -> str:
        """Add a content slide."""
        return await self.call_tool("add_slide", {"title": title, "bullets": bullets})
    
    async def add_image_slide(self, title: str, image_prompt: str = "", caption: str = "") -> str:
        """Add an image slide."""
        return await self.call_tool("add_image_slide", {
            "title": title,
            "image_prompt": image_prompt,
            "caption": caption,
        })
    
    async def save_presentation(self, filename: str = "output.pptx") -> str:
        """Save the presentation."""
        return await self.call_tool("save_presentation", {"filename": filename})
    
    async def plan_slides(self, topic: str, num_slides: int = 5) -> list[dict]:
        """Generate a slide plan."""
        raw = await self.call_tool("plan_slides", {"topic": topic, "num_slides": num_slides})
        parsed = _extract_json(raw)
        
        if isinstance(parsed, list) and parsed:
            result = []
            for item in parsed:
                if isinstance(item, dict) and "title" in item:
                    result.append({
                        "title": str(item["title"]),
                        "bullets": [str(b) for b in item.get("bullets", [])],
                    })
                elif isinstance(item, str):
                    result.append({"title": item, "bullets": []})
            if result:
                return result
        
        # Fallback
        try:
            from utils.planner import build_fallback_plan
            return build_fallback_plan(topic, num_slides)
        except ImportError:
            # Simple fallback
            return [{"title": topic, "bullets": []}] + [
                {"title": f"Slide {i+1}", "bullets": ["Key point 1", "Key point 2", "Key point 3"]}
                for i in range(num_slides - 1)
            ]
    
    async def generate_content(self, title: str) -> list[str]:
        """Generate content for a slide."""
        raw = await self.call_tool("generate_content", {"title": title})
        parsed = _extract_json(raw)
        
        if isinstance(parsed, list) and parsed:
            return [str(b) for b in parsed if str(b).strip()]
        
        # Fallback
        if isinstance(raw, str) and raw.strip():
            lines = [
                l.strip().lstrip("-•*1234567890.) ").strip()
                for l in raw.splitlines() if l.strip()
            ]
            lines = [l for l in lines if len(l) > 3]
            if lines:
                return lines[:5]
        
        return ["Key concept overview", "Important details", "Real-world application", "Summary"]