import csv
 
 # go thru all the emails in the full email roster
with open('all_names.csv') as all_names:
    reader = csv.DictReader(all_names)
    for row in reader:
        email = row['Email']
        name_found = False
        
        # check to see if the email is in the names we want to delete csv
        with open('names_to_delete.csv') as names_to_delete:
            second_reader = csv.reader(names_to_delete)
            for entry in second_reader:
                if email == entry[0]:
                    name_found = True
        
        # if the name is not one we want to delete, put it in the new csv
        if not name_found:
            with open('new_names.csv', 'a', newline='') as new_names:
                writer = csv.writer(new_names)
                first_name = row['First Name']
                last_name = row['Last Name']
                full_row = [first_name, last_name, email]
                writer.writerow(full_row)