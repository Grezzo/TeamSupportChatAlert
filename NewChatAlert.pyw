# TODO: Get Cookie from other browsers (IE and Firefox). See
# https://bitbucket.org/richardpenman/browser_cookie (and perhaps
# contribute)?

# TODO: Raise exceptions better in my functions

import logging
from os import getenv, path
from sqlite3 import connect
from time import sleep, ctime
from ctypes import windll

from win32crypt import CryptUnprotectData
from requests import post, RequestException


class GetChromeCookieError(Exception):
    pass


class GetChatRequestCountError(Exception):
    pass


# Function that displays a message box
def display_message_box(title, text, style):
    windll.user32.MessageBoxW(0, text, title, style)


# Function that returns session cookie from chrome
def get_chrome_cookie_value(name):
    # Connect to Chrome's cookies db
    cookies_database_path = getenv(
        "APPDATA") + r"\..\Local\Google\Chrome\User Data\Default\Cookies"
    if not path.exists(cookies_database_path):
        logging.warning("Cookies database missing")
        display_message_box("Chrome cookies database missing",
               "It looks like you're not using Chrome.\n\n"
               "This application only works if you use Chrome",
               16)
        exit(1)
    conn = connect(cookies_database_path)
    cursor = conn.cursor()
    # Get the encrypted cookie
    cursor.execute(
        "SELECT encrypted_value FROM cookies WHERE name IS '" + name + "'")
    results = cursor.fetchone()
    # Close db
    conn.close()
    if results is None:
        decrypted = None
    else:
        decrypted = CryptUnprotectData(results[0], None, None, None, 0)[
            1].decode("utf-8")
    return decrypted


# Function that returns chat status using a provided session cookie
def get_chat_request_count(cookie):
    # Ask TeamSupport for the chat status using cookie
    try:
        response = post(
            "https://app.teamsupport.com/chatstatus",
            cookies={"TeamSupport_Session": cookie},
            data='{"lastChatMessageID": -1, "lastChatRequestID": -1}'
            )
    except RequestException as e:
        chat_request_count = None
    else:
        try:
            chat_request_count = response.json()["ChatRequestCount"]
        except ValueError as e:
            chat_request_count = None
    return chat_request_count


def main():
    # Loop forever - checking for new chat requests
    while True:
        cookie = get_chrome_cookie_value("TeamSupport_Session")
        if cookie is None:
            logging.warning("Cookie not found")
            display_message_box(
                "Session cookie not found",
                "TeamSupport session cookie could not be found in Chrome "
                "store\n\n"
                "Log in to TeamSupport using Chrome to fix this",
                16)
            # Pause for 30 seconds before trying again
            sleep(30)
        else:
            chat_request_count = get_chat_request_count(cookie)
            if chat_request_count is None:
                logging.warning("Web request failed")
                display_message_box("Web Request Failed",
                       "Chat request count could not be retrieved from "
                       "TeamSupport\n\n"
                       "New chat notifications will not work until this is "
                       "resolved\n\n"
                       "Log in to TeamSupport using Chrome to fix this",
                       16)
                # Pause for 30 seconds before trying again
                sleep(30)
            else:
                # Alert if there are new chat requests or log if none
                if chat_request_count == 0:
                    logging.info("No new chat requests")
                elif chat_request_count == 1:
                    logging.info("1 new chat request")
                    display_message_box(
                        "New Chat Request",
                        "There is 1 new chat request",
                        64)
                else:
                    logging.info(
                        str(chat_request_count) +
                        " new chat requests")
                    display_message_box("New Chat Requests", "There are " +
                           str(chat_request_count) + " chat requests", 48)
                # Pause for 10 seconds before checking again
                sleep(10)

if __name__ == "__main__":
#    raise GetChromeCookieError('Test')
    # Set up logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s: %(message)s')
    # Run script
    main()
