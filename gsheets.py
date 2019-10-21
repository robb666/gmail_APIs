from oauth2client.service_account import ServiceAccountCredentials
import gspread


scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('creds.json', scope)
client = gspread.authorize(creds)
sheet = client.open('Test2').sheet1

def save_spreadsheet(nazwa, jaka, telefonytekst, mailetekst, ileocen, wspolne_slowa):
    insertRow = [nazwa, jaka, telefonytekst, mailetekst, ileocen, wspolne_slowa]
    sheet.insert_row(insertRow, 2)

