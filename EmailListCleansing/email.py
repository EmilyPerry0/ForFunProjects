import win32com.client
import csv

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")


# open the csv file
file = open("names_to_delete.csv", 'w', newline='')
writer = csv.writer(file)

# 60 -> 83 (substring of emails)
for x in range(150):
    try:
        # get the email and put it in a csv file
        msg = outlook.OpenSharedItem(r"D:\Documents-D\\Non-GitHub_Coding\\Oulook parsing\\Undeliverable_ _EXTERNAL_ Delivery Status Notification (Failure) (" + str(x) + ").msg")
        
        # find the end of the email addy
        final_index = msg.Body.index("because") - 1

        # put it in the csv
        email = [msg.Body[60:final_index]]
        writer.writerow(email)
        del msg
    except:
        # having some errors is okay! Because for some reason I did not actually recieve all the emails 1-150.
        print("error!")

# 136 -> 152 (substring of emails)
for x in range(149):
    try:
        # get the email and put it in a csv file
        msg = outlook.OpenSharedItem(r"D:\Documents-D\\Non-GitHub_Coding\\Oulook parsing\\Undeliverable_ Astronomy Club Newsletter 9_15 (" + str(x) + ").msg")

        # find the end of the email addy
        final_index = msg.Body.index("couldn't") - 1
        email = [msg.Body[136:final_index]]
        writer.writerow(email)
        del msg
    except:
        # having some errors is okay! Because for some reason I did not actually recieve all the emails 1-149.
        print("error! 2: electric boogaloo")

# cleanup
del outlook
file.close()