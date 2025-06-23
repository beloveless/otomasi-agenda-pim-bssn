from telegram import Bot

TOKEN = '7135848774:AAGlJLq4eR4WXaS6k8KHL4RoJkywJ09b7mo'  # ganti dengan token bot Anda

bot = Bot(token=TOKEN)

updates = bot.get_updates()
for update in updates:
    print(update.message.chat.id, update.message.from_user.username)
