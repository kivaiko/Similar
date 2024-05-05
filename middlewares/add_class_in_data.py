from typing import Any, Awaitable, Callable, Dict
from aiogram import BaseMiddleware
from aiogram.types import Update
from telethon import TelegramClient


class AddClassInDataMiddleware(BaseMiddleware):
    def __init__(self, client: TelegramClient) -> None:
        self.client = client

    async def __call__(
        self,
        handler: Callable[[Update, Dict[str, Any]], Awaitable[Any]],
        event: Update,
        data: Dict[str, Any]
    ) -> Any:
        data["client"] = self.client
        return await handler(event, data)