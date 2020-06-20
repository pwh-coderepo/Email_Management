#!/usr/local/bin/python3.7
# encoding: utf-8
"""
CBExchange -- read/delete emails based on position/mailId in the mailbox. Return email body
@author:    Shanto Mathew
@copyright:  @2020
@contact:    smathew@paperlesswarehousing.com.au
"""
dfdfdfd

import sys
from exchangelib import Credentials, Account

__all__ = []
__version__ = 0.1
__date__ = '2020-06-11'
__updated__ = '2020-06-19'
TESTRUN = True
LIMIT = 10
USERNAME, PASSWORD = "support@paperlesswarehousing.com.au", "Paperless01"
POS_ADJUSTMENT = True  # Allow position number passed to start with 1 rather than 0


class Exchange():
    def __init__(self, username, password):
        """
        Class initiation. It logs in to the MS Exchange mailbox.
        :param username: username to the mailbox
        :param password: password to the mailbox
        """

        self.username = username
        self.password = password
        self.account = username
        self.mailbox = self.login()

    def login(self):
        """
        This function logs the user into the MS Exchange mailbox
        :return: The mailbox account object
        """
        credentials = Credentials(self.username, self.password)
        return Account(self.account, credentials=credentials, autodiscover=True)

    def get_email_body(self, position):
        """
        Gets the email body of the email that is at the position passed to this function
        :param position: Position of the email in the mailbox
        :return: Email Body
        """
        email = list(self.mailbox.inbox.all().order_by('-datetime_received')[:LIMIT])[position]
        _email_body = email.text_body
        _email_id = email.id
        _email_subject = email.subject

        return _email_subject, _email_body, _email_id

    def delete_email(self, _email_id):
        """
        This function deletes the email that has the emailId passed to it (emailId here is not the same as the email
        address)
        :param _email_id: A long alphanumeric string which is the ID of one single email in the mailbox.
        :return: True/False, Message depending on the outcome of the delete action
        """
        all_email_list = self.mailbox.inbox.all().order_by('-datetime_received')[:LIMIT + 5]
        delete_email_list = [email for email in all_email_list if email.id == _email_id]

        if len(delete_email_list) != 1:
            return False, "More or less than one email found for deletion. With subjects: {}. No email deleted".format(
                ", ".join([x.subject for x in delete_email_list]))
        else:
            for email in delete_email_list:
                return_value = email.delete()
                if return_value is None:
                    return True, "Email deleted with subject: {}".format(email.subject)
                else:
                    return False, "Failed to delete email with subject: {}".format(email.subject)


def validate_arguments(_args):
    """
    This function validates the input arguments
    :param _args: arguments passed to the script starting from position 1
    :return: True/False, Message depending on the outcome of the validation
    """

    _args[1] = str(_args[1])

    if len(_args) != 2:
        return False, "Please pass exactly 2 arguments"
    elif _args[0] not in ["-r", "-d"]:
        return False, "Please pass command argument. -r for reading email and -d for deleting email"
    elif _args[0] == "-r" and (not _args[1].isnumeric() or int(_args[1]) < 0):
        return False, "Please pass a positive number as the second argument corresponding to the position of email in " \
                      "the mailbox "
    elif _args[0] == "-d" and len(_args[1]) < 100:
        return False, "The emailId looks inaccurate. Please check that and try again. It usually is an alphanumeric " \
                      "string with length of around 151 characters "
    else:
        return True, "Validation passed"


if __name__ == "__main__":
    if __name__ == "__main__":
        try:
            if TESTRUN:
                # args = ["-r", 1]
                args = ["-d", "AAMkAGQ5YWRlMGM5LTI0MGEtNDJjYS05ODMzLWZkYzhhOWVlYTkxYQBGAAAAAACWis8uQyOEQ6Oh81UzsRPSBwDA3bYyBaXeR7N4oHIE1yL1AAAAAAEMAADA3bYyBaXeR7N4oHIE1yL1AAADdLqbAAA="]

            else:
                args = sys.argv[1:]
            valid, message = validate_arguments(args)
            if valid:
                account = Exchange(USERNAME, PASSWORD)
                # Read Email and return emailId
                if args[0] == "-r":
                    # Allow position number passed to start with 1 rather than 0
                    if POS_ADJUSTMENT:
                        args[1] = int(args[1]) - 1
                    email_subject, email_body, email_id = account.get_email_body(int(args[1]))
                    print(email_subject, email_body, email_id)
                # Delete Email
                elif args[0] == "-d":
                    response, message = account.delete_email(args[1])
                    print(response, message)
            else:
                print("Validation of arguments Failed with message: {}".format(message))
        except Exception as e:
            import traceback

            print("Script finished with Error/Exception")
            traceback.print_exc()


        else:
            print("Script finished without Error/Exception")
        finally:
            pass
            # account.protocol.close()
