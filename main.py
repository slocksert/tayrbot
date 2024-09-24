from tayrbot import TayrBot

try:
    tayrbot = TayrBot()

    email_input, pwd_input, login_btn = tayrbot.find_login_inputs()
    tayrbot.interact_with_login_inputs(email_input, pwd_input, login_btn)

    tayrbot.close_ad()
    tayrbot.find_menu_icon_and_move_through()
    tayrbot.switch_to_iframe_and_submit()

    tayrbot.run()
    tayrbot.driver.quit()
except KeyboardInterrupt:
    tayrbot.driver.quit()
    quit()
except Exception as e:
    print(e)