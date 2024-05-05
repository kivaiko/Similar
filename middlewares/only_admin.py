from aiogram import BaseMiddleware
from typing import Callable, Awaitable, Any, Dict
from aiogram.types import Message
from config import MAIN_ADMIN


class OnlyAdminsMiddleware(BaseMiddleware):
    async def __call__(
        self,
        handler: Callable[[Message, Dict[str, Any]], Awaitable[Any]],
        event: Message, 
        data: Dict[str, Any]
    ) -> Any:
        if event.from_user.id == MAIN_ADMIN:
            return await handler(event, data)
        else:
            return None