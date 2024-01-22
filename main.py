import User__login
from termcolor import colored

username = ""
account_type = ""
# making wb object with sheets
wb, sheet1, sheet2 = User__login.loading_xl_file()

User__login.opening_message()

User__login.xl_defaulter_or_ok()

# Asks Login or Sign-up
login_or_signup = input(colored("Would you like to login or signup?\n", 'green'))
while login_or_signup.lower() != "login" and login_or_signup.lower() != "signup":
    login_or_signup = input(colored("You have inputted invalid answer. Would you like to login or signup?\n", 'red'))

# Login
if login_or_signup.lower() == "login":
    username = input(colored("Username: ", "cyan"))
    password = input(colored('Password: ', 'cyan'))
    username, password, account_type, name, cnic, phone_no = User__login.username(username, password)


# Signup Using Function
elif login_or_signup.lower() == "signup":
    username, password, account_type, name, cnic, phone_no = User__login.signup()

choice = ""
# Account Type is buyer, admin, hr, finance


if account_type.lower() == "buyer":
    while choice != "5":
        choice = User__login.buyer_choice()
        if choice == "1":
            buy_options_full_detail_lst, s_no_value = User__login.show_available_units()
            if buy_options_full_detail_lst is not None:
                # Payment
                User__login.xl_update(str(s_no_value), str(username))

        elif choice == "2":
            display_only = True
            buy_options_full_detail_lst, s_no_value = User__login.bought_units_display(username, True)
        elif choice == "3":
            display_only = False
            buy_options_full_detail_lst, s_no_value = User__login.bought_units_display(username, False)

            if buy_options_full_detail_lst is not None:
                # Payment
                User__login.xl_update(str(s_no_value), str(username))
        elif choice == "4":
            User__login.sending_whatsapp_query(username)


elif account_type.lower() == "hr":
    while choice != "3":
        choice = User__login.hr_choice()
        if choice == "1":
            User__login.question_checker()
        if choice == "2":
            User__login.plaza_update()


elif account_type.lower() == "finance":
    while choice != "4":
        choice = User__login.finance_choice()
        if choice == "1":
            defaulter_name_lst, defaulter_phone_no_lst = User__login.defaulter_display()
        elif choice == "2":
            User__login.monthly_expense()
        elif choice == "3":
            defaulter_name_lst, defaulter_phone_no_lst = User__login.defaulter_display()
            User__login.defaulter_msg(defaulter_name_lst, defaulter_phone_no_lst)


elif account_type.lower() == "admin":
    while choice != "4":
        choice = User__login.admin_choice()
        if choice == "1":  # View Finance Report
            User__login.read_file()
        elif choice == "2":  # Refund Granted
            User__login.refund_granted()
        elif choice == "3":  # Look at Inventory
            User__login.all_units_display()

User__login.goodbye_msg()

