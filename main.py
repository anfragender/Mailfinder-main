import sys
from qtpy import QtWidgets
import socket
import smtplib
import dns.resolver
import csv
import xlsxwriter
from datetime import datetime

from ui.mainwindow import Ui_MainWindow

app = QtWidgets.QApplication(sys.argv)


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setWindowTitle("EmailFinder")

        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        self.readCsvFile("contacts.csv")
        self.ui.entryButton.clicked.connect(self.onNewEntry)
        self.ui.saveButton.clicked.connect(self.onSafe)
        self.ui.actionSave.triggered.connect(self.onSafe)
        self.ui.deleteButton.clicked.connect(self.onDelete)
        self.ui.hinweisLabel.hide()
        self.ui.exportButton.clicked.connect(self.onExport)
        self.ui.actionExport.triggered.connect(self.onExport)

    def readCsvFile(self, filename):
        self.ui.tableWidget.setRowCount(0)
        with open(filename, "r", newline='') as file:
            reader = csv.reader(file, delimiter=',', quotechar='"')
            for line in reader:
                row = self.ui.tableWidget.rowCount()
                self.ui.tableWidget.insertRow(row)
                self.ui.tableWidget.setItem(row, 0, QtWidgets.QTableWidgetItem(line[0]))
                self.ui.tableWidget.setItem(row, 1, QtWidgets.QTableWidgetItem(line[1]))
                self.ui.tableWidget.setItem(row, 2, QtWidgets.QTableWidgetItem(line[2]))
                self.ui.tableWidget.setItem(row, 3, QtWidgets.QTableWidgetItem(line[3]))
                self.ui.tableWidget.setItem(row, 4, QtWidgets.QTableWidgetItem(line[4]))
                self.ui.tableWidget.setItem(row, 5, QtWidgets.QTableWidgetItem(line[5]))
                self.ui.tableWidget.setItem(row, 6, QtWidgets.QTableWidgetItem(line[6]))
                self.ui.tableWidget.setItem(row, 7, QtWidgets.QTableWidgetItem(line[7]))
                self.ui.tableWidget.setItem(row, 8, QtWidgets.QTableWidgetItem(line[8]))
                self.ui.tableWidget.setItem(row, 9, QtWidgets.QTableWidgetItem(line[9]))
                self.ui.tableWidget.setItem(row, 10, QtWidgets.QTableWidgetItem(line[10]))

    def spellchecker(self, text):
        mapping = {ord(u"ß"): u"ss", ord(u"ä"): u"ae", ord(u"ö"): u"oe", ord(u"ü"): u"ue"}
        checked_text = text.translate(mapping)
        return checked_text

    def formats(self, first, last, domain):
        """
        Create a list of 20 possible email formats combining:
        - First name:          [empty] | Full | Initial |
        - Delimitator:         [empty] |   .  |    _    |    -
        - Last name:           [empty] | Full | Initial |
        """
        list = []
        list.append(first + '.' + last + '@' + domain)  # first.last@example.com
        list.append(last + '@' + domain)  # last@example.com
        # list.append(first[0] + '@' + domain)  # f@example.com
        list.append(first[0] + last + '@' + domain)  # flast@example.com
        list.append(first[0] + '.' + last + '@' + domain)  # f.last@example.com
        list.append(first[0] + '_' + last + '@' + domain)  # f_last@example.com
        list.append(first[0] + '-' + last + '@' + domain)  # f-last@example.com
        list.append(first + '@' + domain)  # first@example.com
        list.append(first + last + '@' + domain)  # firstlast@example.com
        list.append(first + '_' + last + '@' + domain)  # first_last@example.com
        list.append(first + '-' + last + '@' + domain)  # first-last@example.com
        # list.append(first[0] + last[0] + '@' + domain)  # fl@example.com
        # list.append(first[0] + '.' + last[0] + '@' + domain)  # f.l@example.com
        # list.append(first[0] + '-' + last[0] + '@' + domain)  # f_l@example.com
        # list.append(first[0] + '-' + last[0] + '@' + domain)  # f-l@example.com
        list.append(first + last[0] + '@' + domain)  # fistl@example.com
        # list.append(first + '.' + last[0] + '@' + domain)  # first.l@example.com
        # list.append(first + '_' + last[0] + '@' + domain)  # fist_l@example.com
        # list.append(first + '-' + last[0] + '@' + domain)  # fist-l@example.com
        # list.append(last[0] + '@' + domain)  # l@example.com

        return list

    def verify(self, list, domain):
        """
        Create a list of all valid addresses out of a list of emails.
        """

        valid = []

        for email in list:
            try:
                records = dns.resolver.resolve(domain, 'MX')
            except (dns.resolver.NXDOMAIN, dns.resolver.NoAnswer):
                print('DNS query could not be performed.')

                # quit()
                return valid

            # Get MX record for the domain
            mx_record = records[0].exchange
            mx = str(mx_record)

            # Get local server hostname
            local_host = socket.gethostname()

            # Connect to SMTP
            smtp_server = smtplib.SMTP()
            smtp_server.connect(mx)
            smtp_server.helo(local_host)
            smtp_server.mail(email)
            code, message = smtp_server.rcpt(email)

            try:
                smtp_server.quit()
            except smtplib.SMTPServerDisconnected:
                print('Server disconnected. Verification could not be performed.')
                # quit()
                return valid
            # Add to valid addresses list if SMTP response is positive
            if code == 250:
                valid.append(email)
            else:
                continue

        return (valid)

    def return_valid(self, valid, possible):
        """
        Return final output comparing list of valid addresses to the possible ones:
        1. No valid  > Return message
        2. One valid > Copy to clipboard
        3. All valid > Catch-all server
        4. Multiple  > List addresses
        """
        leer = []

        #  if len(valid) == 0:
        #      print('No valid email address found')
        #      return valid
        #
        #       elif len(valid) == 1:
        #          print('Valid email address found : ' + valid[0])
        #         return valid
        #
        #       elif len(valid) == len(possible):
        #          print('Catch-all server. Verification not possible.')
        #         return leer
        #    else:
        #       print('Multiple valid email addresses found:')
        #      for address in valid:
        #         print(address)
        #    return valid

        if len(valid) == len(possible):
            print('Catch-all server. Verification not possible.')
            return leer
        else:
            return valid

    def mailfinder(self, firstname, lastname, domain):
        first_name = firstname  # First name
        last_name = lastname  # Last name
        domain_name = domain  # Domain

        emails_list = self.formats(first_name, last_name, domain_name)
        print("Emailliste erstellt")
        valid_list = self.verify(emails_list, domain_name)
        print("Validliste erstellt")
        validmails = []
        if len(valid_list) > 0:
            validmails = self.return_valid(valid_list, emails_list)
        return validmails

    def onDelete(self):
        self.ui.tableWidget.setRowCount(0)
        self.ui.hinweisLabel.show()

    def onNewEntry(self):
        print("Start des Programms")

        # erstmal alles Auslesen
        firma_value = self.ui.entryFirma.text()
        domain_value = self.ui.entryDomain.text()
        position_value = self.ui.entryPosition.text()
        titel_value = self.ui.entryTitel.text()
        vorname_value = str(self.ui.entryVorname.text())
        nachname_value = self.ui.entryNachname.text()
        sprache_value = self.ui.entryLanguage.text()
        # Anrede braucht eine Fallunterscheidung,da mehrere Auswahlen
        anrede_value = ""
        if self.ui.buttonMale.isChecked():
            anrede_value = "Herr"
        elif self.ui.buttonFemale.isChecked():
            anrede_value = "Frau"
        linkedin_value = self.ui.entryLinkedin.text()
        # testen ob überhaupt Daten eingeben wurden, sonst nichts machen
        # if len(vorname_value) >0 & len(nachname_value) >0 & len(domain_value) >0:
        #    return

        # Email: überprüfen ob bereits eine angegeben ist
        email_value = self.ui.entryMail.text()
        checked_firstname = self.spellchecker(vorname_value)
        checked_lastname = self.spellchecker(nachname_value)

        # die Funktion Mailfinder soll nur ausgeführt werden, wenn der Hacken steht & die 3 Daten vorhanden sind
        email_list_value = ""
        if self.ui.checkBoxSearch.isChecked():
            email_list = self.mailfinder(checked_firstname, checked_lastname, domain_value)
            email_list_value = str(email_list)
            if len(email_list) > 0:
                email_value = email_list[0]
                print(email_value)
            else:
                email_value = ""
        else:
            print("CheckBoxSearch not checked")

        # jetzt alle Einträge in die Tabelle Schreiben und anschließend in separater Funktion in CSV speichern?
        row = self.ui.tableWidget.rowCount()
        self.ui.tableWidget.insertRow(row)
        self.ui.tableWidget.setItem(row, 0, QtWidgets.QTableWidgetItem(firma_value))
        self.ui.tableWidget.setItem(row, 1, QtWidgets.QTableWidgetItem(domain_value))
        self.ui.tableWidget.setItem(row, 2, QtWidgets.QTableWidgetItem(position_value))
        self.ui.tableWidget.setItem(row, 3, QtWidgets.QTableWidgetItem(titel_value))
        self.ui.tableWidget.setItem(row, 4, QtWidgets.QTableWidgetItem(anrede_value))
        self.ui.tableWidget.setItem(row, 5, QtWidgets.QTableWidgetItem(vorname_value))
        self.ui.tableWidget.setItem(row, 6, QtWidgets.QTableWidgetItem(nachname_value))
        self.ui.tableWidget.setItem(row, 7, QtWidgets.QTableWidgetItem(sprache_value))
        self.ui.tableWidget.setItem(row, 8, QtWidgets.QTableWidgetItem(email_value))
        self.ui.tableWidget.setItem(row, 9, QtWidgets.QTableWidgetItem(email_list_value))
        self.ui.tableWidget.setItem(row, 10, QtWidgets.QTableWidgetItem(linkedin_value))
        # nachdem alles eingetragen wurde soll lieber zwischengespeichert werden
        self.onSafe()

    def onSafe(self):
        with open('contacts.csv', 'w', newline='') as file:
            writer = csv.writer(file, delimiter=",", quotechar='"')
            rows = self.ui.tableWidget.rowCount()
            for i in range(0, rows):
                rowContent = [
                    self.ui.tableWidget.item(i, 0).text(),
                    self.ui.tableWidget.item(i, 1).text(),
                    self.ui.tableWidget.item(i, 2).text(),
                    self.ui.tableWidget.item(i, 3).text(),
                    self.ui.tableWidget.item(i, 4).text(),
                    self.ui.tableWidget.item(i, 5).text(),
                    self.ui.tableWidget.item(i, 6).text(),
                    self.ui.tableWidget.item(i, 7).text(),
                    self.ui.tableWidget.item(i, 8).text(),
                    self.ui.tableWidget.item(i, 9).text(),
                    self.ui.tableWidget.item(i, 10).text(),
                ]
                writer.writerow(rowContent)
        self.ui.hinweisLabel.hide()

    def onExport(self):

        workbook = xlsxwriter.Workbook('exportContacts' + str(datetime.now().minute) + '.xlsx')
        worksheet = workbook.add_worksheet()

        # widen the first column to make the text clearer
        worksheet.set_column('A:A', 20)
        # Add a bold format to use to highlight cells.
        bold = workbook.add_format({'bold': True})

        # Write Headers
        worksheet.write(0, 0, 'Firma', bold)
        worksheet.write(0, 1, 'Webseite', bold)
        worksheet.write(0, 2, 'Position', bold)
        worksheet.write(0, 3, 'Titel', bold)
        worksheet.write(0, 4, 'Anrede', bold)
        worksheet.write(0, 5, 'Vorname', bold)
        worksheet.write(0, 6, 'Nachname', bold)
        worksheet.write(0, 7, 'Sprache', bold)
        worksheet.write(0, 8, 'Email', bold)
        worksheet.write(0, 9, 'Email-Liste', bold)
        worksheet.write(0, 10, 'Linkedin', bold)

        # Inhalt der Tabelle in Excel-Exportieren
        rows = self.ui.tableWidget.rowCount()
        n = 0
        for i in range(0, rows):
            #Nur Kontakte mit Email-Adressen werden exportiert
            if len(self.ui.tableWidget.item(i, 8).text()) > 1:
                n += 1
                worksheet.write(n, 0, self.ui.tableWidget.item(i, 0).text())
                worksheet.write(n, 1, self.ui.tableWidget.item(i, 1).text())
                worksheet.write(n, 2, self.ui.tableWidget.item(i, 2).text())
                worksheet.write(n, 3, self.ui.tableWidget.item(i, 3).text())
                worksheet.write(n, 4, self.ui.tableWidget.item(i, 4).text())
                worksheet.write(n, 5, self.ui.tableWidget.item(i, 5).text())
                worksheet.write(n, 6, self.ui.tableWidget.item(i, 6).text())
                worksheet.write(n, 7, self.ui.tableWidget.item(i, 7).text())
                worksheet.write(n, 8, self.ui.tableWidget.item(i, 8).text())
                worksheet.write(n, 9, self.ui.tableWidget.item(i, 9).text())
                worksheet.write(n, 10, self.ui.tableWidget.item(i, 10).text())

        workbook.close()


window = MainWindow()

window.show()
sys.exit(app.exec_())
