from aiogram import Router
from aiogram.types import Message
from aiogram.filters.command import CommandStart


router = Router()


@router.message(CommandStart())
async def start_work_bot(message: Message):
    await message.answer("📨 Отправьте файл с расширением .xlsx")