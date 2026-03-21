import PyPDF2
import os 
try:
    from .GetData import GetData, Futam, Horses
except:
    from GetData import GetData, Futam, Horses
import re
import sys
from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, 
                             QPushButton, QDialog, QCalendarWidget, QLabel)
from PyQt6.QtCore import QDate

class DatePickerPopup(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Please select the race date.")
        self.layout = QVBoxLayout(self)

        # The Calendar Widget
        self.calendar = QCalendarWidget(self)
        self.calendar.setGridVisible(True)
        
        # Select Button
        self.select_btn = QPushButton("Confirm Selection")
        self.select_btn.clicked.connect(self.accept)

        self.layout.addWidget(self.calendar)
        self.layout.addWidget(self.select_btn)

    def get_selected_date(self):
        # Returns a QDate object
        return self.calendar.selectedDate()



def removeTXT(search, txt):
    if search in txt:   return txt[0:txt.index(search)]
    else:                 return txt

def remove_dupl(li:list):
    ndli = []
    for ln in li:
        if ln not in ndli:
            ndli.append(ln)
    
    return ndli

def clean_opinion(txt):
    # Find the last occurrence of ')' that belongs to a horse and cut after it
    match = re.search(r'(.*\(\d+\))', txt)
    if match:
        return match.group(1)
    return txt           


class ReadPDF:

    def __init__(self, mainWindow=None, file_name=""):
        self.rome_num = ["I","II","III","IV","V","VI","VII","VIII","IX","X","XI","XII","XIII","XIV"]
        self.horses = []
        self.futams = []
        self.opinions = []
        self.pdf = []
        self.date = []
        self.mainWindow = mainWindow
        self.read(file_name)
        


    def read(self,file_name):
        if not os.path.exists(file_name): 
            print("File not exsits")
            return False
        
        reader = PyPDF2.PdfReader(file_name)
        num_pages = len(reader.pages)


        for page_num in range(num_pages):
            page = reader.pages[page_num]
            text = page.extract_text()
            self.pdf.append(text.split("\n"))

        if("\\" in file_name): date = file_name.replace(".pdf","").split("\\")[-1].split("_")
        else:                  date = file_name.replace(".pdf","").split("/")[-1].split("_")
        to_remove = {"galopp", "versenyprogram"}
        filtered_date = [item for item in date if item not in to_remove]
        #print(filtered_date)
        try:
            y,m,d = filtered_date
        except:
            if self.mainWindow != None:
                popup = DatePickerPopup(self.mainWindow)
                if popup.exec():
                    date = popup.get_selected_date()
                    y,m,d =  date.toString("yyyy-MM-dd").split("-")
            else:
                date = input("pls give me date (yyyy-MM-dd): ").split("-")
                if len(date) == 3: y,m,d = date 
                else: print("faild to load data")



        data = GetData(f"https://mla.kincsempark.hu/racecards/gallop/{y}-{m}-{d}")
        for ln in data.futam_data:
            title = Futam()
            title.load_json(ln)  # or however you load futam
            self.futams.append(title)

            for horse in ln["participants"]:
                jockey = Horses()
                jockey.load_json(horse, title.id)
                self.horses.append(jockey)


        opinions = []
        for page in self.pdf:
            for ln in page:
                for horse in self.horses:
                    if "Véleményünk:" in ln and horse.Hname in ln:
                        opinions.append(ln)
                        break

        for i, op in enumerate(opinions):
            opinions[i] = opinions[i].strip()
            opinions[i] = opinions[i].replace("\n",' ')
            opinions[i] = opinions[i].replace("\t",' ')
            
            #for num in range(9, 14):
                #opinions[i] = removeTXT(str(num), opinions[i])

            opinions[i] = removeTXT("Elérhetőségek", opinions[i])
            opinions[i] = removeTXT("100.000 Ft", opinions[i])
            opinions[i] = removeTXT("200.000 Ft", opinions[i])
            opinions[i] = removeTXT("300.000 Ft", opinions[i])
            opinions[i] = removeTXT("101.190 Ft", opinions[i])
            opinions[i] = removeTXT("Esélyelemzés", opinions[i])
            opinions[i] = removeTXT("Nyeregbe!", opinions[i])
            opinions[i] = removeTXT("Az ügető", opinions[i])
            opinions[i] = removeTXT("A futamban", opinions[i])
            
            opinions[i] = clean_opinion(opinions[i])

            #Véleményünk: 
            opinions[i] = opinions[i].replace("Véleményünk: ",'')
            opinions[i] = opinions[i].replace("Véleményünk:",'')
            opinions[i] = opinions[i].strip()
            
            
            

        
        self.opinions = remove_dupl(opinions)

        """
        print("\nopinions:")
        for i,op in enumerate(self.opinions):
            print(f"{i}: {op}")
        """
        
        
        cnt = 0
        for futam in self.futams:
            if not futam.daily in self.rome_num:
                cnt = futam.id+1
        #print("\nfutam opinions:")
        for i,futam in enumerate(self.futams):

            if futam.daily in self.rome_num:
              
                try:
                    futam.opinion = self.opinions[futam.id-cnt]
                except:
                    print("IndexError: list index out of range\n Opinions missing an element")
                #print(f"{i},{futam.id-cnt}: {futam.opinion}")



if __name__ == "__main__":
    PDF_data = ReadPDF(file_name=r"C:\Users\Becsei Szabolcs\Downloads\versenyprogram_2025_12_31_ugeto.pdf")

    print("titles:")
    for i in PDF_data.futams:
        print(i.opinion)



    



    
