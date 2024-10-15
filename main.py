import os
import sys
import sessionhandler
import portfolioticketshandler
import portfolioshareshandler
import excelexporter

#python -m pip install --index-url=https://blpapi.bloomberg.com/repository/releases/python/simple/ blpapi

current_dir = os.path.dirname(__file__)
data_tickets_dir = os.path.join(current_dir, 'shared-data')

os.makedirs(data_tickets_dir, exist_ok=True)

def start_session():
    session = sessionhandler.start_session()

    if not session:
        print("Failed to start session")
        return False
    
    return session

def retrieve_all_ticket_data(session):
    """
    
    """

    print("Retrieving ticket data.")

    while True:
        division_entries = ['EMC','TCH','GM','FIR','RM']
        print("Input division that the data will be appended to. Allowed entries: ['EMC','TCH','GM','FIR','RM'].")
        division = input("Input: ")
        if division in division_entries:
            break
        else:
            print("Invalid input. Please try again.")

    print("Drag and drop ticket history .csv file.")
    ticket_file_raw = input()

    portfolio_tickets_total = portfolioticketshandler.load_portfolio_tickets_total(ticket_file_raw)
    portfolio_shares_total = portfolioshareshandler.load_portfolio_shares_total(portfolio_tickets_total)

    history_dir = os.path.join(data_tickets_dir, "history", division)
    os.makedirs(history_dir, exist_ok=True)
    excelexporter.export_dict_to_excel(os.path.join(history_dir, "history.xlsx"),portfolio_shares_total)


if __name__ == "__main__":
    session = start_session()
    
    retrieve_all_ticket_data(session)