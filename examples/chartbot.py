#!/usr/bin/env python
# pylint: disable=W0613, C0116
# type: ignore[union-attr]
# This program is dedicated to the public domain under the CC0 license.

"""
Bot that attaches charts from xlsx files.
"""
import os
import json
import logging
from pathlib import Path
from typing import Dict

import tempfile
import excel2img
from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.ext import Updater, CommandHandler, CallbackQueryHandler, CallbackContext

logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO
)
logger = logging.getLogger(__name__)

ROOT = Path(__file__).parent.parent.parent
TOKEN = "TOKEN"
ALLOWED_USER_IDS = [42]


def get_charts_data() -> Dict:
    return json.load(open(ROOT / "charts_data.json", "r", encoding='utf-8'))


def get_chart_image(fn_excel: str, fn_image: str, sheet_name: str, cells_range: str):
    excel2img.export_img(ROOT / fn_excel, fn_image, "", f"{sheet_name}!{cells_range}")


def start(update: Update, context: CallbackContext) -> None:
    user_id = update.effective_user.id
    if user_id not in ALLOWED_USER_IDS:
        print("Unauthorized access denied for {}.".format(user_id))
        update.message.reply_text("User disallowed")
        return

    charts_data = get_charts_data()
    keyboard = [[InlineKeyboardButton(chart_name, callback_data=chart_name)] for chart_name in charts_data.keys()]

    reply_markup = InlineKeyboardMarkup(keyboard)
    update.message.reply_text("Please choose:", reply_markup=reply_markup)


def button(update: Update, context: CallbackContext) -> None:
    query = update.callback_query

    # CallbackQueries need to be answered, even if no notification to the user is needed
    # Some clients may have trouble otherwise. See https://core.telegram.org/bots/api#callbackquery

    query.answer()
    chat_id = query.message.chat.id
    charts_data = get_charts_data()
    chart_data = charts_data[query.data]

    try:
        fn_image = "chart.png"
        get_chart_image(fn_image=fn_image, **chart_data)

        query.bot.send_photo(chat_id, photo=open(fn_image, "rb"))
        os.remove(fn_image)
    except OSError as e:
        print(e)
        update.effective_message.reply_text("The chart you are looking for is missing")


def help_command(update: Update, context: CallbackContext) -> None:
    update.message.reply_text("Use /start to test this bot.")


def main():
    # Create the Updater and pass it your bot's token.
    updater = Updater(TOKEN)

    updater.dispatcher.add_handler(CommandHandler('start', start))
    updater.dispatcher.add_handler(CallbackQueryHandler(button))
    updater.dispatcher.add_handler(CommandHandler('help', help_command))

    # Start the Bot
    updater.start_polling()

    # Run the bot until the user presses Ctrl-C or the process receives SIGINT,
    # SIGTERM or SIGABRT
    updater.idle()


if __name__ == '__main__':
    main()
