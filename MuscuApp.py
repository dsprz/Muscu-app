#TODO si pas PyQt5 et si pas pandas
#TODO delete un exercice en particulier
#from win32com.client import Dispatch #pip install pywin32
import configparser
import sys
from excelManager import ExcelManager
from datetime import date
from PyQt5 import QtGui, QtCore
from PyQt5.QtWidgets import (
    QApplication, 
    QLabel, 
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLineEdit,
    QComboBox,
    QPushButton, 
    QListWidget,
    QGridLayout,
    QScrollArea,
    QGroupBox,
    QFormLayout,
    QMenuBar,
    QMainWindow
    )


### FR
__debutDate__ = "02/06/2023"
__author__ = "Jimmy VU"

configFilePath = r"./config.ini"
exercisesFilePath = r"./exercises.ini"
greenPlusSignPath = r"./images/greenPlusSign.png"
redCrossSignPath = r"./images/redCrossSign.png"
windowIconPath = r"./images/moi.png"
editPenIconPath = r"./images/editPen.png"
excelPath = r"./Sport2.xlsx"

config = configparser.ConfigParser()
config.read(configFilePath)
today = date.today()
currentDay = today.strftime("%d/%m/%Y")

"""xl = Dispatch('Excel.Application')
wb = xl.Workbooks.Add()
wb.Close(True, r'C:\Path\to\folder\Sport2.xlsx')"""

class Interface(QWidget):

    greenPlusButtonIconSize = 48
    editPenButtonIconSize = 58

    def __init__(self):
        self.app = QApplication(sys.argv)
        QWidget.__init__(self)
        self.setWindowTitle("MuscuApp")
        self.setWindowIcon(QtGui.QIcon(windowIconPath))
        self.setGeometry(200, 200, 200, 60) #prendre res user
        """self.verticalLayout = QVBoxLayout()
        self.horizontalLayout1 = QHBoxLayout()
        self.horizontalLayout2 = QHBoxLayout()
        self.horizontalLayout3 = QHBoxLayout()
        self.horizontalLayoutMenu = QHBoxLayout()
        self.setLayout(self.verticalLayout)
        self.verticalLayout.addLayout(self.horizontalLayoutMenu)
        self.verticalLayout.addLayout(self.horizontalLayout1)
        self.verticalLayout.addLayout(self.horizontalLayout2)
        self.verticalLayout.addLayout(self.horizontalLayout3)"""

        #New Session Button
        self.newSessionButton = QPushButton(self)
        self.newSessionButtonIcon = QtGui.QPixmap(greenPlusSignPath)
        self.newSessionButtonIcon.setMask(self.newSessionButtonIcon.createMaskFromColor(QtGui.QColor(255, 0, 0)))
        self.newSessionButtonIcon = QtGui.QIcon(self.newSessionButtonIcon)
        self.newSessionButton.setIcon(self.newSessionButtonIcon)
        self.newSessionButton.setIconSize(QtCore.QSize(self.greenPlusButtonIconSize, self.greenPlusButtonIconSize))
        self.newSessionButton.setStyleSheet("QPushButton{background: transparent;}")
        self.newSessionButton.setGeometry(5, 5, self.greenPlusButtonIconSize, self.greenPlusButtonIconSize)
        self.newSessionButton.setToolTip("Nouvelle séance")
        self.newSessionButton.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.newSessionButton.clicked.connect(self.openSessionWindow)

        self.editSessionButton = QPushButton(self)
        self.editSessionButtonIcon = QtGui.QPixmap(editPenIconPath)
        self.editSessionButtonIcon.setMask(self.editSessionButtonIcon.createMaskFromColor(QtGui.QColor(255, 255, 255)))
        self.editSessionButtonIcon = QtGui.QIcon(self.editSessionButtonIcon)
        self.editSessionButton.setIcon(self.editSessionButtonIcon)
        self.editSessionButton.setIconSize(QtCore.QSize(self.editPenButtonIconSize, self.editPenButtonIconSize))
        self.editSessionButton.setStyleSheet("QPushButton{background: transparent;}")
        self.editSessionButton.setGeometry(65, 1, self.editPenButtonIconSize, self.editPenButtonIconSize)
        self.editSessionButton.setToolTip("Modifier une séance")
        self.editSessionButton.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.editSessionButton.clicked.connect(self.openEditSessionWindow)

        #UserInfos
        #self.userInfos(self.horizontalLayout1)

        #NewSessionInfos
        #self.newSessionInfos(self.horizontalLayout2)
    
    def userInfos(self, layout):
        self.createLabel(layout, "Infos utilisateur")
        self.createLabel(layout, currentDay)
        self.createEntry(layout, "Poids en kg", False, (5, 5))
        self.createEntry(layout, "Temps", False, (5,5))
        self.createComboBox(layout, "Shape", 
                            [str(i) for i in range(1, 11)])
    
    def newSessionInfos(self, layout):
        self.createLabel(layout, "Rentrer sa séance")
        self.createComboBox(layout, "Exos", ["Tractions", 
                                     "Curl", 
                                     "Développé couché",
                                     ])
        self.createEntry(layout, "Séries d'échauffement", False)
        self.createEntry(layout, "Séries")
        self.createEntry(layout,"Poids associé à chaque série")
    
    def modifySession(self, layout):
        self.createLabel(layout, "Modifier une séance")
        self.createComboBox(layout, 
                            "Sélectionnez une séance",
                            ["date1",
                             "date2"]
                             )

    def createLabel(self, layout, text,  organize = False):
        label = QLabel(self)
        label.setText(text)
        if organize:
            layout.addWidget(label, organize)
        else:
            layout.addWidget(label)
    
    def createComboBox(self, layout, currentText, list):
        comboBox = QComboBox(self)
        comboBox.addItems(list)
        comboBox.setEditable(True)
        comboBox.setCurrentIndex(-1)
        comboBox.setCurrentText(currentText)
        layout.addWidget(comboBox)

    def createEntry(self, layout, text, organize = False, resize = False):
        #organize: int
        #resize: tuple
        entry = QLineEdit(self)
        entry.setPlaceholderText(text)
        if organize:
            layout.addWidget(entry, organize)
        else:
            layout.addWidget(entry)
        if resize:
            entry.resize(resize[0], resize[1])
    
    """def createButton(self, 
                     layout = False, 
                     toolTipText = "", 
                     buttonText = False, 
                     imagePath = False, 
                     tupleGeometry = False):
        button = QPushButton(self)
        if buttonText:
            button.setText(buttonText)
        if imagePath:
            icon = QtGui.QPixmap(imagePath)
            icon.setMask(icon.createMaskFromColor(QtGui.QColor(255, 0, 0)))
            buttonIcon = QtGui.QIcon(icon)
            button.setIcon(buttonIcon)
            if tupleGeometry:
                button.setIconSize(QtCore.QSize(tupleGeometry[0], tupleGeometry[1]))
        if layout:
            layout.addWidget(button)
        button.setStyleSheet("QPushButton{background: transparent;}")
        button.setGeometry(0,0, tupleGeometry[0], tupleGeometry[1])
        button.setToolTip(toolTipText)"""

    def openSessionWindow(self):
        self.newWorkingSession = NewWorkoutSessionMainWindow()
        self.newWorkingSession.run()
        #self.newWorkingSession.runWorkingSessionWindow()
    
    def openEditSessionWindow(self):
        self.editSession = NewEditSessionMainWindow()
        self.editSession.run()

    def runApp(self):
        self.show()
        self.app.exec_()
    

class WorkingSession(QWidget):

    buttonIconSize = 32
    seriesButtonIconSize = 12
    addSeriesButtonExerciseNumberIconSize = 26
    exercisesBoxItems = [
        "Dips",
        "Tractions",
        "Curl",
        "Développé couché",
        "Rowing Barre",
        "Abdos"
        ]
    
    def __init__(self):
        QWidget.__init__(self)
        self.setWindowTitle("Nouvelle séance")
        self.setWindowIcon(QtGui.QIcon(windowIconPath))
        self.setGeometry(400, 200, 1080, 640) #prendre res user
        self.setFixedHeight(640)

        self.horizontalLayout1 = QHBoxLayout(self)
        self.setLayout(self.horizontalLayout1)
        self.verticalLayout = QVBoxLayout(self)

        self.formLayout = QFormLayout()

        self.groupBox = QGroupBox()
        self.groupBox.setLayout(self.formLayout)

        self.scrollArea = QScrollArea()  
        self.scrollArea.setWidgetResizable(True)
        self.scrollArea.setWidget(self.groupBox)
        self.verticalLayout.addWidget(self.scrollArea)


        self.createSessionExercisesBoxItems()
        self.horizontalLayout1.addLayout(self.verticalLayout)
        self.exerciseNumber = 0

        self.deleteExerciseButtonIcon = QtGui.QPixmap(redCrossSignPath)
        self.deleteExerciseButtonIcon.setMask(self.deleteExerciseButtonIcon.createMaskFromColor(QtGui.QColor(0, 255, 0)))
        self.deleteExerciseButtonIcon = QtGui.QIcon(self.deleteExerciseButtonIcon)
    
        self.deleteSeriesButtonIcon = QtGui.QPixmap(redCrossSignPath)
        self.deleteSeriesButtonIcon.setMask(self.deleteSeriesButtonIcon.createMaskFromColor(QtGui.QColor(0, 255, 0)))
        self.deleteSeriesButtonIcon = QtGui.QIcon(self.deleteExerciseButtonIcon)

        self.addSeriesButtonExerciseNumberIcon = QtGui.QPixmap(greenPlusSignPath)
        self.addSeriesButtonExerciseNumberIcon.setMask(self.addSeriesButtonExerciseNumberIcon.createMaskFromColor(QtGui.QColor(255, 0, 0)))
        self.addSeriesButtonExerciseNumberIcon = QtGui.QIcon(self.addSeriesButtonExerciseNumberIcon)

        self.exercice = []
        self.allExercises = []
        #self.scrollArea.setWidget(QWidget(self.scrollArea))
        #self.verticalLayout.addWidget(self.scrollArea)


    def createSessionExercisesBoxItems(self):

        #emptylabel pour contourner le problème de la listbox qui ne veut pas s'alginer à gauche
        self.emptyLabel = QLabel(self)
        sizeToBeAdded = 50
        self.exercisesBoxItems = sorted(self.exercisesBoxItems)
        self.exercises = QListWidget()
        for index, value in enumerate(self.exercisesBoxItems):
            self.exercises.insertItem(index, value)
        self.exercises.setMaximumWidth(self.exercises.sizeHintForColumn(0) + sizeToBeAdded)
        self.horizontalLayout1.addWidget(self.exercises)
        self.horizontalLayout1.addWidget(self.emptyLabel, 0, QtCore.Qt.AlignLeft)
        self.exercises.itemClicked.connect(self.showExerciseInfos)

    def showExerciseInfos(self, clickedItem):
        """Crée et affiche tout un exercice"""
        self.exerciseNumber+=1
        self.gridLayout = QGridLayout(self)
        self.gridLayout.setVerticalSpacing(10)
        nb = self.exerciseNumber

        #cheat pour contourner un problème
        try:
            self.horizontalLayout1.removeWidget(self.emptyLabel)
            self.emptyLabel.deleteLater()
        except Exception:
            pass

        """setattr(self,f"gridLayout{self.exerciseNumber}", QGridLayout(self))
        setattr(self,f"exerciseNumber{self.exerciseNumber}", QLabel(self))
        setattr(self,f"repetitionEntry{self.exerciseNumber}", QLineEdit(self))
        setattr(self,f"massEntry{self.exerciseNumber}", QLineEdit(self))
        setattr(self,f"deleteExerciseButton{self.exerciseNumber}", QPushButton(self))"""

        self.labelFont = QtGui.QFont("Poppins", 15)
        self.labelFont.setBold(True)
        self.label = QLabel(self)
        self.label.setText(clickedItem.text().upper())
        self.label.setFont(self.labelFont)
        self.label.setAlignment(QtCore.Qt.AlignCenter)
        self.label.setStyleSheet("border: 1px solid black;")
        self.label.setFixedHeight(50)


        self.series = []
        numberOfEntries = len(config["DEFAULT"])
        self.gridLayout.addWidget(self.label, 0, 0, 2, numberOfEntries + 1)

        self.addSeriesButtonExerciseNumber = QPushButton(self)
        self.addSeriesButtonExerciseNumber.setIcon(self.addSeriesButtonExerciseNumberIcon)
        self.addSeriesButtonExerciseNumber.setIconSize(QtCore.QSize(self.addSeriesButtonExerciseNumberIconSize, 
                                                                    self.addSeriesButtonExerciseNumberIconSize))
        self.addSeriesButtonExerciseNumber.setStyleSheet("QPushButton{background: transparent;}")
        self.addSeriesButtonExerciseNumber.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.addSeriesButtonExerciseNumber.clicked.connect(self.addSeries)

        row = 1
        while row <= int(config["DEFAULT"]["Series"]):
            #création d'une série
            self.exercice.append(self.createSeries(self.gridLayout, row))
            row+=1

        self.allExercises.append(self.exercice)
        #self.exercice = [ [[label1, series1, repet1, lbs1], [label1, series 2, repet1, lbs1], etc... ]], ...]
        self.exercice=[]

        self.addSeriesButtonExerciseNumber.setToolTip("Ajouter une série pour cet exercice")
        self.addSeriesButtonExerciseNumber.clicked.connect(self.addSeries)

        setattr(self,f"deleteExerciseButton{nb}", QPushButton(self))
        getattr(self,f"deleteExerciseButton{nb}").setIcon(self.deleteExerciseButtonIcon)
        getattr(self,f"deleteExerciseButton{nb}").setIconSize(QtCore.QSize(self.buttonIconSize, 
                                                                           self.buttonIconSize))
        getattr(self,f"deleteExerciseButton{nb}").setStyleSheet("QPushButton{background: transparent;}")
        getattr(self,f"deleteExerciseButton{nb}").setGeometry(5, 5, self.buttonIconSize, self.buttonIconSize)
        getattr(self,f"deleteExerciseButton{nb}").setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        deleteText = f"Supprimer l'exercice {nb}"
        getattr(self,f"deleteExerciseButton{nb}").setToolTip(deleteText)
        getattr(self,f"deleteExerciseButton{nb}").clicked.connect(lambda : self.removeExercise(nb))

        setattr(self,f"addSeriesButtonExerciseNumber{nb}", QPushButton(self))
        getattr(self,f"addSeriesButtonExerciseNumber{nb}").setIcon(self.addSeriesButtonExerciseNumberIcon)
        getattr(self,f"addSeriesButtonExerciseNumber{nb}").setIconSize(QtCore.QSize(self.addSeriesButtonExerciseNumberIconSize, 
                                                                                    self.addSeriesButtonExerciseNumberIconSize))
        getattr(self,f"addSeriesButtonExerciseNumber{nb}").setStyleSheet("QPushButton{background: transparent;}")
        #getattr(self,f"addSeriesButtonExerciseNumber{nb}").setGeometry(5, 5, self.buttonIconSize, self.buttonIconSize)
        getattr(self,f"addSeriesButtonExerciseNumber{nb}").setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        addSeriesText = f"Ajouter une {5}ème série"
        getattr(self,f"addSeriesButtonExerciseNumber{nb}").setToolTip(addSeriesText)
        getattr(self,f"addSeriesButtonExerciseNumber{nb}").clicked.connect(lambda : self.addSeries(nb))

        self.gridLayout.addWidget(getattr(self,f"addSeriesButtonExerciseNumber{nb}"), row + 1, 0, 1, 1)
        self.gridLayout.addWidget(getattr(self,f"deleteExerciseButton{nb}"), 0, 3, 2, 1)

        self.formLayout.addRow(self.gridLayout)

        self.verticalLayout.setAlignment(QtCore.Qt.AlignTop)

        #autoscrolldown
        self.scrollArea.verticalScrollBar().rangeChanged.connect(lambda: self.scrollArea.verticalScrollBar().setValue(9999))
    

    def removeExercise(self, exerciseNumber):
        """Enlève un exercice déjà ajouté"""
        print(f"argument = {exerciseNumber}")
        #try:
        layoutToBeCleared = self.allExercises[exerciseNumber-1][-1][-1]
        #except IndexError: #si je sélectionne le tout dernier exercice directement
            #layoutToBeCleared = self.allExercises[exerciseNumber-2][-1][-1]

        self.formLayout.takeRow(exerciseNumber-1)
        for i in reversed(range(layoutToBeCleared.count())):
            #layoutToBeCleared.itemAt(i).widget().setContentsMargins(0, 0, 0, 0)
            layoutToBeCleared.itemAt(i).widget().setParent(None)
            #layoutToBeCleared.removeWidget(layoutToBeCleared.itemAt(i).widget())
            #layoutToBeCleared.itemAt(i).widget().deleteLater()"""

        for j in range(exerciseNumber, len(self.allExercises)):
            getattr(self,f"deleteExerciseButton{j+1}").setParent(None)
            self.allExercises[j][-1][-1].addWidget(getattr(self,f"deleteExerciseButton{j}"), 0, 3, 2, 1)

        #print(f"je vais enlever le {exerciseNumber-1}e indice alors qu'il y en a {len(self.allExercises)-1}")

        self.allExercises.remove(self.allExercises[exerciseNumber-1])
        self.exerciseNumber-=1

    def removeSeries(self, series):
        pass

    def createSeries(self, layout, row):
        """Crée une row-ème série d'un exercice"""
        column = 0
        self.seriesEntry = QLabel(self)
        self.repetitionsEntry = QLineEdit(self)
        self.massEntry = QLineEdit(self)
        self.deleteSeriesButton = QPushButton(self)

            #self.deleteExerciseButton = QPushButton(self)

        self.seriesEntry.setText(f"Série {row}")
        self.seriesEntry.setStyleSheet("border: 1px solid black;")
        self.seriesEntry.setMinimumWidth(150)
        self.seriesEntry.setAlignment(QtCore.Qt.AlignCenter)

        self.repetitionsEntry.setText(config["DEFAULT"]["Repetitions"])
        self.massEntry.setText(config["DEFAULT"]["Mass"])

        self.deleteSeriesButton.setIcon(self.deleteSeriesButtonIcon)
        self.deleteSeriesButton.setIconSize(QtCore.QSize(self.seriesButtonIconSize, self.seriesButtonIconSize))
        self.deleteSeriesButton.setStyleSheet("QPushButton{background: transparent;}")
        self.deleteSeriesButton.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        deleteText = f"Supprimer la série {row}"
        self.deleteSeriesButton.setToolTip(deleteText)
        self.deleteSeriesButton.clicked.connect(lambda : self.removeSeries(row))

        layout.addWidget(self.seriesEntry, row + 1, column, 1, 1)
        column+=1
        layout.addWidget(self.repetitionsEntry, row + 1, column, 1, 1)
        column+=1
        layout.addWidget(self.massEntry, row + 1, column, 1, 1)
        column+=1
        layout.addWidget(self.deleteSeriesButton, row + 1, column, 1, 1)

        self.series = [self.label,
                        self.seriesEntry, 
                        self.repetitionsEntry, 
                        self.massEntry,
                        self.addSeriesButtonExerciseNumber,
                        self.deleteSeriesButton, 
                        #self.deleteExerciseButton, 
                        layout
                        ]
        return self.series
        
    def addSeries(self, exerciseNumber):
        """Ajoute une série à l'exercice exerciceNumber"""

        currentSeriesNumber = len(self.allExercises[exerciseNumber-1])
        newSeriesNumber = currentSeriesNumber + 1
        rowForNewButton = newSeriesNumber + 1 

        print(f"je print {exerciseNumber}")
        layout = self.allExercises[exerciseNumber-1][-1][-1]
        self.allExercises[exerciseNumber-1].append(self.createSeries(layout, newSeriesNumber))
       

        #Bouge le bouton pour ajouter une série une row de plus
        try:
            getattr(self, f"addSeriesButtonExerciseNumber{exerciseNumber}").deleteLater()
        except RuntimeError:
            pass
        setattr(self,f"addSeriesButtonExerciseNumber{exerciseNumber}", QPushButton(self))
        getattr(self,f"addSeriesButtonExerciseNumber{exerciseNumber}").setIcon(self.addSeriesButtonExerciseNumberIcon)
        getattr(self,f"addSeriesButtonExerciseNumber{exerciseNumber}").setIconSize(QtCore.QSize(self.addSeriesButtonExerciseNumberIconSize, self.addSeriesButtonExerciseNumberIconSize))
        getattr(self,f"addSeriesButtonExerciseNumber{exerciseNumber}").setStyleSheet("QPushButton{background: transparent;}")
        getattr(self,f"addSeriesButtonExerciseNumber{exerciseNumber}").setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        addSeriesText = f"Ajouter une {newSeriesNumber + 1}ème série"
        getattr(self,f"addSeriesButtonExerciseNumber{exerciseNumber}").setToolTip(addSeriesText)
        getattr(self,f"addSeriesButtonExerciseNumber{exerciseNumber}").clicked.connect(lambda : self.addSeries(exerciseNumber))

        layout.addWidget(getattr(self,f"addSeriesButtonExerciseNumber{exerciseNumber}"), rowForNewButton + 1, 0, 1, 1)
    
    def sendToExcel(self):
        """Envoie tous les exercices sur Excel"""
        excelManager = ExcelManager()
        #mergerIndex = 1
        #row = 1
        dataframe = []
        seriesNumber = int(config["DEFAULT"]["Series"])

        for exercise in self.allExercises:
            for series in exercise:
                print(series)
                usefulData = [series[i].text() for i in range(seriesNumber)]
                for index, data in enumerate(usefulData):
                    try:
                        usefulData[index] = int(data)
                    except Exception:
                        pass
                dataframe.append(usefulData)
        print(dataframe)
        excelManager.writeInExcel(dataframe, self.exerciseNumber)
        self.clearAllExercises()

    def printExcelDatabase(self):
        excelManager = ExcelManager()
        #print(excelManager.returnExcelDataframe())
        excelManager.testToExcel()

    def clearAllExercises(self):
        """Efface tous les exercices de l'interface"""
        for exercise in self.allExercises:
            row=0
            layout = exercise[-1][-1]
            for i in reversed(range(layout.count())):
            #layoutToBeCleared.itemAt(i).widget().setContentsMargins(0, 0, 0, 0)
                layout.itemAt(i).widget().deleteLater()
            self.formLayout.takeRow(row)         
        self.allExercises = []
        self.exerciseNumber = 0

    def runWorkingSessionWindow(self):
        self.show()

class EditSession(QWidget):
    buttonIconSize = 32
    seriesButtonIconSize = 12
    addSeriesButtonExerciseNumberIconSize = 26
    def __init__(self):
        QWidget.__init__(self)
        self.setWindowTitle("Nouvelle séance")
        self.setWindowIcon(QtGui.QIcon(windowIconPath))
        self.setGeometry(400, 200, 1080, 640) #prendre res user
        self.setFixedHeight(640)

        self.horizontalLayout1 = QHBoxLayout(self)
        self.setLayout(self.horizontalLayout1)
        self.verticalLayout = QVBoxLayout(self)

        self.formLayout = QFormLayout()

        self.groupBox = QGroupBox()
        self.groupBox.setLayout(self.formLayout)

        self.scrollArea = QScrollArea()  
        self.scrollArea.setWidgetResizable(True)
        self.scrollArea.setWidget(self.groupBox)
        self.verticalLayout.addWidget(self.scrollArea)

        self.deleteExerciseButtonIcon = QtGui.QPixmap(redCrossSignPath)
        self.deleteExerciseButtonIcon.setMask(self.deleteExerciseButtonIcon.createMaskFromColor(QtGui.QColor(0, 255, 0)))
        self.deleteExerciseButtonIcon = QtGui.QIcon(self.deleteExerciseButtonIcon)

        self.deleteSeriesButtonIcon = QtGui.QPixmap(redCrossSignPath)
        self.deleteSeriesButtonIcon.setMask(self.deleteSeriesButtonIcon.createMaskFromColor(QtGui.QColor(0, 255, 0)))
        self.deleteSeriesButtonIcon = QtGui.QIcon(self.deleteSeriesButtonIcon)


        self.addSeriesButtonExerciseNumberIcon = QtGui.QPixmap(greenPlusSignPath)
        self.addSeriesButtonExerciseNumberIcon.setMask(self.addSeriesButtonExerciseNumberIcon.createMaskFromColor(QtGui.QColor(255, 0, 0)))
        self.addSeriesButtonExerciseNumberIcon = QtGui.QIcon(self.addSeriesButtonExerciseNumberIcon)

        self.excelManager = ExcelManager()
        self.createSessionsList()
        self.horizontalLayout1.addLayout(self.verticalLayout)
        self.exerciseNumber = 0
        self.printExcelDatabase()
        self.exercise = []
        self.allExercises = []


    def printExcelDatabase(self):
        #print(excelManager.returnExcelDataframe())
        self.excelManager.testToExcel()

    def createSessionsList(self):
        """Crée la QComboBox qui contient toutes les sessions"""
        self.excelManagerDateList = self.excelManager.getDateList()
        allSessionsList = [self.excelManagerDateList[i] for i in range(len(self.excelManagerDateList)-1) if i%2==0]
        self.emptyLabel = QLabel(self)
        sizeToBeAdded = 50
        self.allSessionsItems = sorted(allSessionsList)
        self.sessions = QListWidget()
        for index, value in enumerate(allSessionsList):
            self.sessions.insertItem(index, value)
        self.sessions.setMaximumWidth(self.sessions.sizeHintForColumn(0) + sizeToBeAdded)
        self.horizontalLayout1.addWidget(self.sessions)
        self.horizontalLayout1.addWidget(self.emptyLabel, 0, QtCore.Qt.AlignLeft)
        self.sessions.itemClicked.connect(lambda: self.showSessionInfos(self.sessions.currentRow()))
    
    def showSessionInfos(self, index):
        """Affiche tous les exercices d'une séance"""
        #print(exercise)
        self.currentSession = index
        print(self.currentSession)
        self.clearWindow()
        self.allSessions = self.excelManager.getWorkoutSessionList()
        self.workoutSessionList = self.allSessions[index]
        for exercise in self.workoutSessionList:
            self.exerciseNumber+=1
            self.gridLayout = QGridLayout(self)
            self.gridLayout.setVerticalSpacing(10)
            exerciseName = exercise[0]
            numberOfSeries = len(exercise[1])
            row = 1
            while row <= numberOfSeries:
                self.exercise.append(self.createSeries(self.gridLayout, row, exerciseName, exercise))
                row+=1
            setattr(self,f"deleteExerciseButton{self.exerciseNumber}", QPushButton(self))
            getattr(self,f"deleteExerciseButton{self.exerciseNumber}").setIcon(self.deleteExerciseButtonIcon)
            getattr(self,f"deleteExerciseButton{self.exerciseNumber}").setIconSize(QtCore.QSize(self.buttonIconSize, 
                                                                            self.buttonIconSize))
            getattr(self,f"deleteExerciseButton{self.exerciseNumber}").setStyleSheet("QPushButton{background: transparent;}")
            getattr(self,f"deleteExerciseButton{self.exerciseNumber}").setGeometry(5, 5, self.buttonIconSize, self.buttonIconSize)
            getattr(self,f"deleteExerciseButton{self.exerciseNumber}").setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
            deleteText = f"Supprimer l'exercice {self.exerciseNumber}"
            getattr(self,f"deleteExerciseButton{self.exerciseNumber}").setToolTip(deleteText)

            setattr(self,f"addSeriesButtonExerciseNumber{self.exerciseNumber}", QPushButton(self))
            getattr(self,f"addSeriesButtonExerciseNumber{self.exerciseNumber}").setIcon(self.addSeriesButtonExerciseNumberIcon)
            getattr(self,f"addSeriesButtonExerciseNumber{self.exerciseNumber}").setIconSize(QtCore.QSize(self.addSeriesButtonExerciseNumberIconSize, 
                                                                                        self.addSeriesButtonExerciseNumberIconSize))
            getattr(self,f"addSeriesButtonExerciseNumber{self.exerciseNumber}").setStyleSheet("QPushButton{background: transparent;}")
            #getattr(self,f"addSeriesButtonExerciseNumber{nb}").setGeometry(5, 5, self.buttonIconSize, self.buttonIconSize)
            getattr(self,f"addSeriesButtonExerciseNumber{self.exerciseNumber}").setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
            addSeriesText = f"Ajouter une {5}ème série"
            getattr(self,f"addSeriesButtonExerciseNumber{self.exerciseNumber}").setToolTip(addSeriesText)
            getattr(self,f"addSeriesButtonExerciseNumber{self.exerciseNumber}").clicked.connect(lambda : self.addSeries(self.exerciseNumber))
            self.allExercises.append(self.exercise)
            self.exercise = []
            self.gridLayout.addWidget(getattr(self,f"deleteExerciseButton{self.exerciseNumber}"), 0, 3, 2, 1)
            self.formLayout.addRow(self.gridLayout)


            #getattr(self,f"deleteExerciseButton{exercise}").clicked.connect(lambda : self.removeExercise(nb))



        #layout.addWidget(getattr(self,f"addSeriesButtonExerciseNumber{exerciseNumber}"), rowForNewButton + 1, 0, 1, 1)
        self.verticalLayout.setAlignment(QtCore.Qt.AlignTop)
    
    def createSeries(self, layout, row, exerciseName, exercise):
        """Crée une row-ème série d'un exercice"""

        numberOfEntries = len(config["DEFAULT"])
        column = 0
        self.seriesEntry = QLabel(self)
        self.repetitionsEntry = QLineEdit(self)
        self.massEntry = QLineEdit(self)
        self.deleteSeriesButton = QPushButton(self)

        self.labelFont = QtGui.QFont("Poppins", 15)
        self.labelFont.setBold(True)
        self.label = QLabel(self)
        self.label.setText(exerciseName.upper())
        self.label.setFont(self.labelFont)
        self.label.setAlignment(QtCore.Qt.AlignCenter)
        self.label.setStyleSheet("border: 1px solid black;")
        self.label.setFixedHeight(50)
        self.gridLayout.addWidget(self.label, 0, 0, 2, numberOfEntries + 2)

            #self.deleteExerciseButton = QPushButton(self)

        self.seriesEntry.setText(f"Série {row}")
        self.seriesEntry.setStyleSheet("border: 1px solid black;")
        self.seriesEntry.setMinimumWidth(150)
        self.seriesEntry.setAlignment(QtCore.Qt.AlignCenter)


        repetitionsList = exercise[2]
        massList = exercise[3]
        repetitionsText = str(int(repetitionsList[row-1]))
        massesText = str(float(massList[row-1]))
        self.repetitionsEntry.setText(repetitionsText)
        self.massEntry.setText(massesText)

        self.deleteSeriesButton.setIcon(self.deleteSeriesButtonIcon)
        self.deleteSeriesButton.setIconSize(QtCore.QSize(self.seriesButtonIconSize, self.seriesButtonIconSize))
        self.deleteSeriesButton.setStyleSheet("QPushButton{background: transparent;}")
        self.deleteSeriesButton.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        deleteText = f"Supprimer la série {row}"
        self.deleteSeriesButton.setToolTip(deleteText)
        #self.deleteSeriesButton.clicked.connect(lambda : self.removeSeries(row))"""

        layout.addWidget(self.seriesEntry, row + 1, column, 1, 1)
        column+=1
        layout.addWidget(self.repetitionsEntry, row + 1, column, 1, 1)
        column+=1
        layout.addWidget(self.massEntry, row + 1, column, 1, 1)
        column+=1
        layout.addWidget(self.deleteSeriesButton, row + 1, column, 1, 1)

        self.series = [self.label,
                        self.seriesEntry, 
                        self.repetitionsEntry, 
                        self.massEntry,
                        #self.addSeriesButtonExerciseNumber,
                        self.deleteSeriesButton, 
                        #self.deleteExerciseButton, 
                        layout
                        ]
        return self.series
    
    def addSeries(self, exerciseNumber):
        """Ajoute une série à l'exercice exerciceNumber"""

        currentSeriesNumber = len(self.allExercises[exerciseNumber-1])
        newSeriesNumber = currentSeriesNumber + 1
        rowForNewButton = newSeriesNumber + 1 

        print(f"je print {exerciseNumber}")
        layout = self.allExercises[exerciseNumber-1][-1][-1]
        self.allExercises[exerciseNumber-1].append(self.createSeries(layout, newSeriesNumber))
       

        #Bouge le bouton pour ajouter une série une row de plus
        try:
            getattr(self, f"addSeriesButtonExerciseNumber{exerciseNumber}").deleteLater()
        except RuntimeError:
            pass
        setattr(self,f"addSeriesButtonExerciseNumber{exerciseNumber}", QPushButton(self))
        getattr(self,f"addSeriesButtonExerciseNumber{exerciseNumber}").setIcon(self.addSeriesButtonExerciseNumberIcon)
        getattr(self,f"addSeriesButtonExerciseNumber{exerciseNumber}").setIconSize(QtCore.QSize(self.addSeriesButtonExerciseNumberIconSize, 
                                                                                                self.addSeriesButtonExerciseNumberIconSize))
        getattr(self,f"addSeriesButtonExerciseNumber{exerciseNumber}").setStyleSheet("QPushButton{background: transparent;}")
        getattr(self,f"addSeriesButtonExerciseNumber{exerciseNumber}").setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        addSeriesText = f"Ajouter une {newSeriesNumber + 1}ème série"
        getattr(self,f"addSeriesButtonExerciseNumber{exerciseNumber}").setToolTip(addSeriesText)
        getattr(self,f"addSeriesButtonExerciseNumber{exerciseNumber}").clicked.connect(lambda : self.addSeries(exerciseNumber))

        layout.addWidget(getattr(self,f"addSeriesButtonExerciseNumber{exerciseNumber}"), rowForNewButton + 1, 0, 1, 1)

    def clearWindow(self):
        """Efface la fenêtre"""
        for exercise in self.allExercises:
            row=0
            layout = exercise[-1][-1]
            for i in reversed(range(layout.count())):
                layout.itemAt(i).widget().deleteLater()
            self.formLayout.takeRow(row)
        self.allExercises = []
        self.exerciseNumber = 0


    def updateSession(self, index):
        #print(self.allSessions[index])
        allRepetitions = []
        allMasses = []
        for exercises in self.allExercises:
           # print(exercises)
           # print("############################")
            intermediateRepetitionsList = []
            intermediateMassesList = []
            for series in exercises:
                repetitions = float(series[2].text())
                mass = float(series[3].text())
                intermediateRepetitionsList.append(repetitions)
                intermediateMassesList.append(mass)
            allRepetitions.append(intermediateRepetitionsList)
            allMasses.append(intermediateMassesList)
        
        sessionIndex = 0
        for sessions in self.allSessions[index]:
            print(sessions)
            sessions[2] = allRepetitions[sessionIndex]
            sessions[3] = allMasses[sessionIndex]
            sessionIndex+=1
        print(self.allSessions[index])

        #print(allMasses)
        #self.allSessions[index][0][2] = allRepetiions
        #self.allSessions[index][0][3] = allMasses
        #self.allSessions[index][0] = self.allSessions[index][0]
        #print(self.allSessions[index][0])


            
    
    def formatAllSessions(self):
        self.allExercisesfromAllSessions = []
        intermediateSessionList = []
        self.numberOfExercisesNumber = []
        exerciseNumber = 0
        for session in self.allSessions:
            #print(session)
            #print("#####################")
            for exercise in session:
                exerciseName = exercise[0]
                seriesNumber = len(exercise[1])
                intermediateList = []
                exerciseNumber+=1
                for series in range(seriesNumber):
                    intermediateList.append(exerciseName)
                    intermediateList.append(exercise[1][series])
                    intermediateList.append(exercise[2][series])
                    intermediateList.append(exercise[3][series])
                    intermediateSessionList.append(intermediateList)
                    intermediateList = []
            self.numberOfExercisesNumber.append(exerciseNumber)
            exerciseNumber = 0
            self.allExercisesfromAllSessions.append(intermediateSessionList)
            intermediateSessionList = []
        #print(self.allExercisesfromAllSessions)

    def sendToExcel(self, index):
        print(index)
        if index == None:
            return
        self.updateSession(index)
        self.formatAllSessions()
        excelManager = ExcelManager()
        dataframe = []
        
        for index, dataframe in enumerate(self.allExercisesfromAllSessions):
            excelManager.writeInExcel(dataframe, self.numberOfExercisesNumber[index], "w" if index == 0 else "a")
        
    def getCurrentSession(self):
        return self.currentSession

    def run(self):
        self.show()

class NewWorkoutSessionMainWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.newSeance = WorkingSession()
        self.setGeometry(200, 200, 1080, 640)
        self.setCentralWidget(self.newSeance)
        self.setWindowTitle("Nouvelle séance")
        self.setWindowIcon(QtGui.QIcon(windowIconPath))
        self.menuBar = QMenuBar(self)
        self.setMenuBar(self.menuBar)
        self.file = self.menuBar.addMenu("Fichier")
        self.file.addAction("Envoyer vers Excel", self.newSeance.sendToExcel)
        self.menuBar.addMenu("Infos")
        
    def run(self):
        self.show()
    
class NewEditSessionMainWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.newEditSeanceWindow = EditSession()
        self.setGeometry(200, 200, 1080, 640)
        self.setCentralWidget(self.newEditSeanceWindow)
        self.setWindowTitle("Modifier une séance")
        self.setWindowIcon(QtGui.QIcon(windowIconPath))
        self.menuBar = QMenuBar(self)
        self.setMenuBar(self.menuBar)
        self.file = self.menuBar.addMenu("Fichier")
        self.file.addAction("Envoyer vers Excel", lambda: self.newEditSeanceWindow.sendToExcel(self.newEditSeanceWindow.getCurrentSession()))
        self.menuBar.addMenu("Infos")

    def run(self):
        self.show()   

if __name__ == "__main__":
    MuscuApp = Interface()
    MuscuApp.runApp()