import os
import sys
import webbrowser
from datetime import datetime
from time import sleep
from PyPDF2 import PdfWriter
from PyQt6.QtCore import QThread, Qt, pyqtSignal
from PyQt6.QtGui import QAction, QIcon, QKeySequence
from PyQt6.QtWidgets import (
	QApplication,
	QMainWindow,
	QVBoxLayout,
	QHBoxLayout,
	QLabel,
	QLineEdit,
	QPushButton,
	QListWidget,
	QWidget,
	QFileDialog,
	QMessageBox,
	QProgressBar,
	QTabWidget
)

# Try to display icon on Windows taskbar
try:
    from ctypes import windll  # Only exists on Windows.
    myappid = 'domnht.pdf.tools.beta-1'
    windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except ImportError:
    pass

class AppConfigs(object):
	# App
	APP_TITLE = "PDF Tools"
	APP_VERSION = "1.0.0 beta"
	APP_AUTHOR = "Nguyen Hieu Thanh"
	# Menu
	MENU_FILE_CAPTION = "File"
	MENU_OPEN_FOLDER_CAPTION = "Open folder"
	MENU_REFRESH_FILES_LIST_CAPTION = "Refresh"
	MENU_SAVE_CAPTION = "Save"
	MENU_CLOSE_CAPTION = "Close"
	MENU_HELP_CAPTION = "Help"
	MENU_HOW_TO_USE_CAPTION = "How to use"
	MENU_ABOUT_CAPTION = "About"
	# UI
	UI_SOURCE_DIRECTORY_PLACEHOLDER = "Select source directory…"
	UI_FILES_LIST_CAPTION = "Files list"
	UI_DOCUMENT_PROPERTIES_CAPTION = "Document properties"
	UI_MERGED_DOCUMENT_NAME = "Merged document"
	# Messages
	DIRECTORY_NOT_VALID_MESSAGE = "Directory is not valid!"
	DIRECTORY_EMPTY_MESSAGE = "Directory contains no PDF files!"
	DOCUMENT_SAVED_SUCCESSFULLY_MESSAGE = "Document saved successfully."
	# Meta fields
	AUTHOR_CAPTION = "Author"
	TITLE_CAPTION = "Title"
	SUBJECT_CAPTION = "Subject"
	KEYWORDS_CAPTION = "Keywords"
	PRODUCER_CAPTION = "Producer"

	def __init__(self, useVietnamese = False):
		super().__init__()
		self.APP_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
		if not useVietnamese: return
		# App
		# self.APP_TITLE = "Gộp tài liệu PDF"
		# self.APP_VERSION = "1.0.0 thử nghiệm"
		self.APP_AUTHOR = "Nguyễn Hiếu Thanh"
		# Menu
		self.MENU_FILE_CAPTION = "Tài liệu"
		self.MENU_OPEN_FOLDER_CAPTION = "Mở thư mục"
		self.MENU_REFRESH_FILES_LIST_CAPTION = "Làm mới"
		self.MENU_SAVE_CAPTION = "Lưu tài liệu"
		self.MENU_CLOSE_CAPTION = "Đóng"
		self.MENU_HELP_CAPTION = "Trợ giúp"
		self.MENU_HOW_TO_USE_CAPTION = "Hướng dẫn"
		self.MENU_ABOUT_CAPTION = "Giới thiệu"
		# UI
		self.UI_SOURCE_DIRECTORY_PLACEHOLDER = "Chọn thư mục nguồn…"
		self.UI_FILES_LIST_CAPTION = "Danh sách tài liệu"
		self.UI_DOCUMENT_PROPERTIES_CAPTION = "Thuộc tính tài liệu"
		self.UI_MERGED_DOCUMENT_NAME = "Tài liệu gộp"
		# Messages
		self.DIRECTORY_ERROR_MESSAGE = "Đường dẫn không hợp lệ!"
		self.DIRECTORY_EMPTY_MESSAGE = "Đường dẫn được chọn không chứa tệp PDF nào!"
		self.FILE_SAVE_SUCCESSFULLY_MESSAGE = "Đã hoàn thành lưu tài liệu."
		# Meta fields
		self.AUTHOR_CAPTION = "Tác giả"
		self.TITLE_CAPTION = "Tên tài liệu"
		self.SUBJECT_CAPTION = "Chủ đề"
		self.KEYWORDS_CAPTION = "Từ khóa"
		self.PRODUCER_CAPTION = "Ứng dụng"

class MainWindow(QMainWindow):
	def __init__(self, appConfigs = None):
		print("MainWindow()")
		super(MainWindow, self).__init__()

		if appConfigs is None: self.configs = AppConfigs()
		else: self.configs = appConfigs

		self.setWindowTitle(self.configs.APP_TITLE)

		# Create Menu bar
		menu = self.menuBar()
		menu.setContentsMargins(0, 0, 0, 0)
		menu.setNativeMenuBar(True)

		self.fileMenu = menu.addMenu(self.configs.MENU_FILE_CAPTION)

		self.pathSelectAction = QAction(
			QIcon(os.path.join(self.configs.APP_BASE_DIR, "icons", "folder-horizontal-open.png")),
			self.configs.MENU_OPEN_FOLDER_CAPTION
		)
		self.pathSelectAction.setShortcut(QKeySequence("Ctrl+O"))
		self.pathSelectAction.triggered.connect(self.selectDirectory)
		self.fileMenu.addAction(self.pathSelectAction)

		self.saveAction = QAction(
			QIcon(os.path.join(self.configs.APP_BASE_DIR, "icons", "disk.png")),
			self.configs.MENU_SAVE_CAPTION
		)
		self.saveAction.setShortcut(QKeySequence("Ctrl+S"))
		self.saveAction.triggered.connect(self.saveFile)
		self.fileMenu.addAction(self.saveAction)

		self.fileMenu.addSeparator()

		self.exitAction = QAction(self.configs.MENU_CLOSE_CAPTION)
		self.exitAction.setShortcut(QKeySequence("Ctrl+W"))
		self.exitAction.triggered.connect(self.close)
		self.fileMenu.addAction(self.exitAction)

		self.helpMenu = menu.addMenu(self.configs.MENU_HELP_CAPTION)

		self.helpAction = QAction(
			QIcon(os.path.join(self.configs.APP_BASE_DIR, "icons", "book-question.png")),
			self.configs.MENU_HOW_TO_USE_CAPTION,
			self
		)
		self.helpAction.triggered.connect(self.showHelp)
		self.helpAction.setEnabled(False)
		self.helpMenu.addAction(self.helpAction)

		self.aboutAction = QAction(
			QIcon(os.path.join(self.configs.APP_BASE_DIR, "icons", "information-white.png")),
			self.configs.MENU_ABOUT_CAPTION, self
		)
		self.aboutAction.triggered.connect(self.aboutThisApp)
		self.aboutAction.setMenuRole(QAction.MenuRole.AboutRole)
		self.aboutAction.setEnabled(False)
		self.helpMenu.addAction(self.aboutAction)

		# Create central widget to contain all elements
		# centralWidget: mainLayout
		#     titleLayout
		#         titleLabel
		#         saveButton
		#     mergerLayout
		#         pathLayout
		#             pathInput
		#             pathRefreshButton
		#             pathSelectButton
		#         tabs
		#             filesList
		#             documentPropertiesWidget: documentPropertiesLayout
		#         progressBar

		centralWidget = QWidget()

		mainLayout = QVBoxLayout()
		mainLayout.setContentsMargins(10, 0, 10, 10)
		mainLayout.setSpacing(5)
		centralWidget.setLayout(mainLayout)

		# Title Layout
		titleLayout = QHBoxLayout()
		mainLayout.addLayout(titleLayout)

		titleLabel = QLabel(self.configs.APP_TITLE)
		font = titleLabel.font()
		font.setPointSize(int(font.pointSize() * 2))
		titleLabel.setFont(font)
		titleLabel.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
		titleLabel.setMinimumHeight(int(titleLabel.sizeHint().height() * 1.2))
		titleLayout.addWidget(titleLabel)

		self.saveButton = QPushButton(
			QIcon(os.path.join(self.configs.APP_BASE_DIR, "icons", "disk.png")),
			self.configs.MENU_SAVE_CAPTION
		)
		self.saveButton.setText("")
		self.saveButton.clicked.connect(self.saveFile)
		self.saveButton.setMaximumWidth(self.saveButton.sizeHint().width())
		titleLayout.addWidget(self.saveButton)

		# Merger Layout
		mergerLayout = QVBoxLayout()
		mergerLayout.setContentsMargins(0, 0, 0, 0)
		mainLayout.addLayout(mergerLayout)

		# Directory Input Layout
		directoryInputLayout = QHBoxLayout()
		directoryInputLayout.setSpacing(5)
		mergerLayout.addLayout(directoryInputLayout)

		self.pathInput = QLineEdit()
		self.pathInput.textChanged.connect(self.listFiles)
		self.pathInput.setReadOnly(True)
		self.pathInput.setPlaceholderText(self.configs.UI_SOURCE_DIRECTORY_PLACEHOLDER)
		directoryInputLayout.addWidget(self.pathInput)

		self.pathRefreshButton = QPushButton(
			QIcon(os.path.join(self.configs.APP_BASE_DIR, "icons", "arrow-circle.png")),
			self.configs.MENU_REFRESH_FILES_LIST_CAPTION
		)
		self.pathRefreshButton.clicked.connect(self.listFiles)
		self.pathRefreshButton.setEnabled(False)
		self.pathRefreshButton.setText("")
		directoryInputLayout.addWidget(self.pathRefreshButton)

		self.pathSelectButton = QPushButton(
			QIcon(os.path.join(self.configs.APP_BASE_DIR, "icons", "folder-horizontal-open.png")),
			self.configs.MENU_OPEN_FOLDER_CAPTION
		)
		self.pathSelectButton.clicked.connect(self.selectDirectory)
		self.pathSelectButton.setText("")
		directoryInputLayout.addWidget(self.pathSelectButton)

		# Files list and Document properties layout
		self.tabs = QTabWidget()
		self.tabs.setAutoFillBackground(True)
		self.tabs.setTabPosition(QTabWidget.TabPosition.North)
		self.tabs.setVisible(False)
		mergerLayout.addWidget(self.tabs)

		# Files list
		self.filesList = QListWidget()
		self.tabs.addTab(self.filesList, self.configs.UI_FILES_LIST_CAPTION)

		# Document properties
		documentPropertiesWidget = QWidget()
		self.tabs.addTab(documentPropertiesWidget, self.configs.UI_DOCUMENT_PROPERTIES_CAPTION)

		documentPropertiesLayout = QVBoxLayout()
		documentPropertiesLayout.setContentsMargins(0, 0, 0, 5)
		documentPropertiesLayout.setSpacing(8)
		documentPropertiesWidget.setLayout(documentPropertiesLayout)

		# Author
		documentAuthorLayout = QVBoxLayout()
		documentAuthorLayout.setContentsMargins(0, 0, 0, 0)
		documentAuthorLayout.setSpacing(3)
		documentPropertiesLayout.addLayout(documentAuthorLayout)

		documentAuthorLabel = QLabel(self.configs.AUTHOR_CAPTION + ":")
		documentAuthorLayout.addWidget(documentAuthorLabel)

		self.documentAuthorInput = QLineEdit()
		self.documentAuthorInput.setPlaceholderText(self.configs.AUTHOR_CAPTION)
		documentAuthorLayout.addWidget(self.documentAuthorInput)

		# Title
		documentTitleLayout = QVBoxLayout()
		documentTitleLayout.setContentsMargins(0, 0, 0, 0)
		documentTitleLayout.setSpacing(3)
		documentPropertiesLayout.addLayout(documentTitleLayout)

		documentTitleLabel = QLabel(self.configs.TITLE_CAPTION + ":")
		documentTitleLayout.addWidget(documentTitleLabel)

		self.documentTitleInput = QLineEdit()
		self.documentTitleInput.setPlaceholderText(self.configs.TITLE_CAPTION)
		documentTitleLayout.addWidget(self.documentTitleInput)

		# Subject
		documentSubjectLayout = QVBoxLayout()
		documentSubjectLayout.setContentsMargins(0, 0, 0, 0)
		documentSubjectLayout.setSpacing(3)
		documentPropertiesLayout.addLayout(documentSubjectLayout)

		self.documentSubjectLabel = QLabel(self.configs.SUBJECT_CAPTION + ":")
		documentSubjectLayout.addWidget(self.documentSubjectLabel)

		self.documentSubjectInput = QLineEdit()
		self.documentSubjectInput.setPlaceholderText(self.configs.SUBJECT_CAPTION)
		documentSubjectLayout.addWidget(self.documentSubjectInput)

		# Keywords
		documentKeywordsLayout = QVBoxLayout()
		documentKeywordsLayout.setContentsMargins(0, 0, 0, 0)
		documentKeywordsLayout.setSpacing(3)
		documentPropertiesLayout.addLayout(documentKeywordsLayout)

		self.documentKeywordsLabel = QLabel(self.configs.KEYWORDS_CAPTION + ":")
		documentKeywordsLayout.addWidget(self.documentKeywordsLabel)

		self.documentKeywordsInput = QLineEdit()
		self.documentKeywordsInput.setPlaceholderText(self.configs.KEYWORDS_CAPTION)
		documentKeywordsLayout.addWidget(self.documentKeywordsInput)

		# Progress Bar
		self.progressBar = QProgressBar()
		self.progressBar.setMinimum(0)
		self.progressBar.setValue(0)
		self.progressBar.setVisible(False)
		mergerLayout.addWidget(self.progressBar)

		# Copyright
		copyrightLabel = QLabel("© " + self.configs.APP_AUTHOR + " 2024.")
		copyrightLabel.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignBottom)
		copyrightLabel.setMinimumHeight(25)
		font = copyrightLabel.font()
		font.setPointSize(int(font.pointSize() * 0.95))
		copyrightLabel.setFont(font)
		mainLayout.addWidget(copyrightLabel)

		# Set main layout
		self.setCentralWidget(centralWidget)
		self.setReadyToSave(False)
		self.setMinimumWidth(400)
		self.setMinimumHeight(self.sizeHint().height())

	def showHelp(self):
		webbrowser.open("https://github.com/domnht/pdf-tools/blob/main/how-to-use.md")

	def aboutThisApp(self):
		webbrowser.open("https://github.com/domnht/pdf-tools/blob/main/readme.md")

	def close(self):
		QApplication.closeAllWindows()

	def selectDirectory(self):
		print("selectDirectory()")
		# fileDialog = QFileDialog(self)
		# fileDialog.setFileMode(QFileDialog.FileMode.Directory)
		selectedDirectory = QFileDialog.getExistingDirectory(self, self.configs.APP_TITLE, os.getcwd())
		print("    selectedDirectory: \"{}\"".format(selectedDirectory))
		if selectedDirectory == "": return
		self.pathInput.setText(selectedDirectory)
		self.pathRefreshButton.setEnabled(True)

	def listFiles(self):
		print("listFiles()")
		self.pathRefreshButton.setEnabled(False)
		selectedDirectory = self.pathInput.text()
		if not os.path.isdir(selectedDirectory):
			self.setReadyToSave(False)
			QMessageBox.critical(self, self.configs.APP_TITLE, self.configs.DIRECTORY_NOT_VALID_MESSAGE)
			return

		currentDirectory = os.getcwd()
		os.chdir(selectedDirectory)
		files = [f for f in os.listdir() if f.upper().endswith(".PDF")]
		files.sort()

		self.filesList.clear()
		for file in files:
			self.filesList.addItem(file)
			print("    file: \"{}\"".format(file))

		if len(files) > 0: self.setReadyToSave()
		else:
			QMessageBox.warning(self, self.configs.APP_TITLE, self.configs.DIRECTORY_EMPTY_MESSAGE)
			self.setReadyToSave(False)

		self.pathRefreshButton.setEnabled(True)
		os.chdir(currentDirectory)

	def saveFile(self):
		print("saveFile()")
		documentAuthor = processInputString(self.documentAuthorInput.text())
		documentTitle = processInputString(self.documentTitleInput.text())
		documentSubject = processInputString(self.documentSubjectInput.text())
		documentKeywords = processInputString(self.documentKeywordsInput.text(), ",;")
		documentProducer = self.configs.APP_TITLE + " " + self.configs.APP_VERSION

		self.documentAuthorInput.setText(documentAuthor)
		self.documentTitleInput.setText(documentTitle)
		self.documentSubjectInput.setText(documentSubject)
		self.documentKeywordsInput.setText(documentKeywords)

		selectedDirectory = self.pathInput.text()
		if not os.path.isdir(selectedDirectory):
			self.disableSaveAction()
			QMessageBox.critical(self, self.configs.APP_TITLE, self.configs.DIRECTORY_NOT_VALID_MESSAGE)
			return

		currentDirectory = os.getcwd()
		os.chdir(selectedDirectory)

		defaultFileName = os.path.join(selectedDirectory, self.configs.UI_MERGED_DOCUMENT_NAME + " " + datetime.now().strftime("%Y%m%d_%H%M%S") + '.pdf')
		targetFile, _ = QFileDialog.getSaveFileName(self, self.configs.APP_TITLE, defaultFileName, "Portable Document Format (*.pdf)")

		if targetFile:
			print("    targetFile: \"{}\"".format(targetFile))

			documentMetadata = {
				'/Author': documentAuthor,
				'/Title': documentTitle,
				'/Subject': documentSubject,
				'/Keywords': documentKeywords,
				'/Producer': documentProducer
			}

			numOfFiles = self.filesList.count()
			self.progressBar.setMaximum(numOfFiles)
			self.progressBar.setVisible(True)
			self.progressBar.setValue(0)
			self.setEnabled(False)

			files = []
			for index in range(numOfFiles): files.append(os.path.join(selectedDirectory, self.filesList.item(index).text()))
			os.chdir(currentDirectory)

			self.thread = MergeThread(files, targetFile, documentMetadata)
			self.thread.progressUpdated.connect(self.setSaveProgress)
			self.thread.start()

	def setSaveProgress(self, value):
		print("setProgressBarValue(value: {})".format(value))
		self.progressBar.setValue(value)
		if value >= self.progressBar.maximum():
			self.setEnabled(True)
			self.progressBar.setVisible(False)
			QMessageBox.information(self, self.configs.APP_TITLE, self.configs.DOCUMENT_SAVED_SUCCESSFULLY_MESSAGE)

	def setReadyToSave(self, isReady = True):
		print("setReadyToSave(isReady: {})".format(isReady))
		if not self.tabs.isVisible(): self.tabs.setVisible(isReady)
		if self.tabs.isVisible(): self.setMinimumHeight(self.sizeHint().height())
		self.saveAction.setEnabled(isReady)
		self.saveButton.setEnabled(isReady)
		self.progressBar.setVisible(False)

	def setEnabled(self, isEnabled = True):
		print("setEnabled(isEnabled: {})".format(isEnabled))
		self.saveButton.setEnabled(isEnabled)
		self.saveAction.setEnabled(isEnabled)
		self.pathSelectButton.setEnabled(isEnabled)
		self.pathSelectAction.setEnabled(isEnabled)

class MergeThread(QThread):
	progressUpdated = pyqtSignal(int)

	def __init__(self, files, targetFile, documentMetadata = {}):
		print("MergeThread(files: list({}), targetFile: \"{}\")".format(len(files), targetFile))
		super(MergeThread, self).__init__()
		self.files = files
		self.targetFile = targetFile
		self.documentMetadata = documentMetadata

	def run(self):
		print("run()")
		pdfMerger = PdfWriter()
		index = 0
		for file in self.files:
			try:
				print("    [{}] {}".format(index, file))
				index += 1
				pdfMerger.append(file)
			except Exception as exception:
				print("    Exception occurred:", exception)
			finally:
				self.progressUpdated.emit(index)
				sleep(0.2)

		try:
			
			pdfMerger.add_metadata(self.documentMetadata)
			pdfMerger.write(self.targetFile)
		except Exception as exception:
			print("    Exception occurred:", exception)
		finally:
			pdfMerger.close()
			self.exit(0)

def processInputString(string, keepChars = ""):
	# remove all special characters except letters, digits and characters in keepChars
	output = "".join(letter for letter in string if letter.isalnum() or letter == " " or letter in keepChars)
	output = output.strip()

	return output

def main():
	appConfigs = AppConfigs(useVietnamese=True)
	os.chdir(appConfigs.APP_BASE_DIR)
	app = QApplication(sys.argv)
	app.setWindowIcon(QIcon(os.path.join(appConfigs.APP_BASE_DIR, "pdf-tools.ico")))
	mainWindow = MainWindow(appConfigs)
	mainWindow.show()

	app.exec()

if __name__ == "__main__": main()
