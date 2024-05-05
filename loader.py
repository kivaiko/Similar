from config import NAME_SESSION, API_ID, API_HASH, BOT_TOKEN, MAIN_ADMIN
from aiogram import Bot, Dispatcher
from aiogram.fsm.storage.memory import MemoryStorage
from telethon import TelegramClient
from middlewares.add_class_in_data import AddClassInDataMiddleware
from middlewares.only_admin import OnlyAdminsMiddleware
from handlers import command_start, excel_files


async def on_startup(bot: Bot):
    try:
        await bot.send_message(MAIN_ADMIN, "✅ Bot turned #on")
    except Exception: pass


async def on_shutdown(bot: Bot):
    try:
        await bot.send_message(MAIN_ADMIN, "⛔️ Bot turned #off")
    except Exception: pass


async def start_bot():
    client = TelegramClient(NAME_SESSION, API_ID, API_HASH, device_model="iPhone 55 Pro",
                            system_version="4.16.30-vxCUSTOM", app_version="1.10.8")
    bot = Bot(BOT_TOKEN, parse_mode="HTML")
    storage = MemoryStorage()
    dp = Dispatcher(storage=storage)
    dp.startup.register(on_startup)
    dp.shutdown.register(on_shutdown)

    dp.update.middleware.register(AddClassInDataMiddleware(client=client))
    dp.message.middleware.register(OnlyAdminsMiddleware())
    dp.include_routers(
        command_start.router,
        excel_files.router
    )

    try:
        await dp.start_polling(bot)
    finally:
        await dp.storage.close()
        await bot.session.close()