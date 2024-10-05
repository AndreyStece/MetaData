import sys, os
import cv2
from PyQt5.QtCore import QDir, QFile
from PyQt5.QtGui import QPixmap, QImage
from PyQt5.QtWidgets import QApplication, QWidget, QTreeView, QFileSystemModel, QMainWindow, QListWidgetItem, QMessageBox
import MetaFind
from metawork import ImageMeta, AudioMeta, VideoMeta, PDFMeta, TextMeta, ArchivesMeta, OfficeMeta, EXEMeta

class MetaDataView(QMainWindow, MetaFind.Ui_MainWindow):
	def __init__(self, dir_path):
		super().__init__()
		# Инициализация ПО
		self.setupUi(self)
		# Запрет на изменение пути в текстовом поле
		self.dir_path.setReadOnly(True)
		# Запрет на изменение элемента в текстовом поле
		self.combo_attr.setReadOnly(True)
		# Масштабирование превью
		self.label_preview.setScaledContents(True)
		# Инициализация списка типов файлов
		list_format = [
			self.tr('all files'),
			self.tr('image'),
			self.tr('audio'),
			self.tr('video'),
			self.tr('pdf'),
			self.tr('text files'),
			self.tr('archives'),
			self.tr('office documents'),
			self.tr('executive files'),
		]
		self.combo_format.addItems(list_format)
		# 1. Окно с директориями
		self.model = QFileSystemModel() # Создаем модель файловой системы
		self.model.setFilter(QDir.Dirs | QDir.Drives | QDir.NoDotAndDotDot | QDir.AllDirs) # Выдаем только директории
		self.model.setRootPath(dir_path) # Задаем корневой каталог
		self.tree_dirs.setModel(self.model) # Задаем модель виджету
		self.tree_dirs.setRootIndex(self.model.index(dirPath)) # Открываем все с текущего индекса
		self.tree_dirs.setColumnWidth(0, 250) # Задаем размер колонок
		self.tree_dirs.setAlternatingRowColors(True) # Задаем альтернативные цвета
		self.tree_dirs.expandAll() # Назначаем дополнительное пространство для виджета
		# 2. Окно с файлами выбранной директории
		self.model = QFileSystemModel() # Создаем модель файловой системы
		self.tree_files.setModel(self.model) # Задаем модель виджету
		self.tree_files.setColumnWidth(0, 250) # Задаем размер колонок
		self.tree_files.setAlternatingRowColors(True) # Задаем альтернативные цвета
		self.tree_files.expandAll() # Назначаем дополнительное пространство для виджета
		self.tree_dirs.selectionModel().selectionChanged.connect(self.onSelectionDir) # Ссылаемся на функцию при выборе директории
		# 3. Выбор типа файлов
		self.combo_format.activated[str].connect(self.onChangeFormat) # Ссылаемся на функцию при выборе типа файлов
		# 4. Выбор файла
		self.tree_files.selectionModel().selectionChanged.connect(lambda: self.onSelectionFile(filePath)) # Ссылаемся на функцию при выборе файла
		# 5. Выбор типа метаданных
		self.combo_meta.activated[str].connect(self.onChangeMeta)
		# 6. Выбор метаданных
		self.meta_data.selectionModel().selectionChanged.connect(self.onChangeAttr)
		# 7. Изменение метаданных
		self.push_edit.clicked.connect(self.onUpdateAttr)
		# 8. Удаление метаданных
		self.push_delItems.clicked.connect(self.onDelAttr)
		# 9. Удаление группы метаданных
		self.push_del.clicked.connect(self.onDelMeta)
		# 10. Сохранение лога метаданных
		self.push_save.clicked.connect(self.onSaveMeta)

	# Функция при выборе директории
	def onSelectionDir(self, *args):
		global filePath
		for sel in self.tree_dirs.selectedIndexes(): # Анализируем только 1-ый столбец
			val = sel.data() # Получаем исходные дочерние данные
			while sel.parent().isValid(): # Следом и получаем остальные родительские
				sel = sel.parent()
				val = sel.data()+ "/" +val
			index=val.find(':') # Преобразуем путь в нормальный вид
			if (index != -1):
				val = val[index-1:].replace(")", "")
			# Модель заново не создаем, так как теряется связь
			self.onShowFiles(val)
			break
		filePath = val
		self.combo_attr.setText('')
		self.new_value.setText('')
		self.meta_data.clear()
		self.combo_meta.clear()
		# Задаем путь в текстовом поле
		self.dir_path.setText(filePath)
	
	# Функция обновления списка файлов
	def onShowFiles(self, path):
		global fileFormat
		self.model.setRootPath(path) # Задаем корневой каталог
		self.model.setFilter(QDir.Files | QDir.NoDotAndDotDot) # Выдаем только файлы
		self.model.setNameFilters(fileFormat) # Выдаем файлы только с определенным набором расширений
		self.tree_files.setModel(self.model) # Задаем модель виджету
		self.tree_files.eventFilter
		self.tree_files.setRootIndex(self.model.index(path)) # Открываем все с текущего индекса

	# Функция при выборе формата файлов
	def onChangeFormat(self, *args):
		global filePath
		global fileFormat
		if (self.combo_format.currentText() == "all files"):
			fileFormat = ["*.*"]
		elif (self.combo_format.currentText() == "image"):
			fileFormat = ["*.jpg", "*.jpeg", "*.png", "*.gif"]
		elif (self.combo_format.currentText() == "audio"):
			fileFormat = ["*.mp3"]
		elif (self.combo_format.currentText() == "video"):
			fileFormat = ["*.mp4", "*.MOV", "*.webm"]
		elif (self.combo_format.currentText() == "pdf"):
			fileFormat = ["*.pdf"]
		elif (self.combo_format.currentText() == "text files"):
			fileFormat = ["*.txt", "*.rtf", "*.py"]
		elif (self.combo_format.currentText() == "archives"):
			fileFormat = ["*.zip", "*.rar", "*.7z"]
		elif (self.combo_format.currentText() == "office documents"):
			fileFormat = ["*.doc", "*.docx", "*.ppt", "*.pptx","*.xsl", "*.xslx","*.csv"]
		elif (self.combo_format.currentText() == "executive files"):
			fileFormat = ["*.exe", "*.dll"]
		self.onShowFiles(filePath)

	# Функция при выборе файла
	def onSelectionFile(self, path):
		global fileName
		global fileExt
		global fileType
		for sel in self.tree_files.selectedIndexes(): # Анализируем только 1-ый столбец
			val = sel.data()
			index=val.find('.')
			fileName = val[:index]
			fileExt = val[index:]
			if (fileExt == ".jpg" or fileExt == ".jpeg" or fileExt == ".png" or fileExt == ".gif"):
				fileType = "image"
				image = QPixmap(filePath + "/" + fileName + fileExt)
				self.label_preview.setPixmap(image)
			elif (fileExt == ".mp3"):
				fileType = "audio"
				self.label_preview.clear()
			elif (fileExt == ".mp4" or fileExt == ".MOV" or fileExt == ".webm"):
				fileType = "video"
				vidcap = cv2.VideoCapture(filePath + "/" + fileName + fileExt)
				success,image = vidcap.read()
				count = 0
				while count != 10 or success != True:
					success,image = vidcap.read()
					count += 1
				if(count != 0 and success == True):
					w, h, ch = image.shape
					qimg = QImage(image.data, h, w, 3*h, QImage.Format_RGB888)
					self.label_preview.setPixmap(QPixmap(qimg))
			elif (fileExt == ".pdf"):
				fileType = "pdf"
				self.label_preview.clear()
			elif (fileExt == ".txt" or fileExt == ".rtf" or fileExt == ".py"):
				fileType = "text files"
				self.label_preview.clear()
			elif (fileExt == ".zip" or fileExt == ".rar" or fileExt == ".7z"):
				fileType = "archives"
				self.label_preview.clear()
			elif (fileExt == ".doc" or fileExt == ".docx" or fileExt == ".ppt" or fileExt == ".pptx" or fileExt == ".xsl" or fileExt == ".xslx"):
				fileType = "office documents"
				self.label_preview.clear()
			elif (fileExt == ".exe" or fileExt == ".dll"):
				fileType = "executive files"
				self.label_preview.clear()
			path = path + '/' + val
			self.combo_attr.setText('')
			self.new_value.setText('')
			self.meta_data.clear()
			self.combo_meta.clear()
			self.onExtractMeta(path)
			break
	
	# Функция при выборе метаданных
	def onChangeAttr(self, *args):
		for k in self.meta_data.selectedIndexes():
			val = k.data()
			if (val.find('#') == -1):
				index = val.find(':')
				val = val[:index]
				self.combo_attr.setText(val)
			else:
				self.combo_attr.setText('')

	# Функция при выборе типа метаданных
	def onChangeMeta(self, *args):
		global metaType
		global metaKey
		global filePath
		global fileName
		global fileExt
		metaType = self.combo_meta.currentText()
		metaKey = True
		self.onExtractMeta(filePath + "/" + fileName + fileExt)

	# Функция вставки списка метаданных
	def onInsertMeta(self):
		global list_meta
		self.combo_meta.clear()
		self.combo_meta.addItems(list_meta)

	# Функция при извлечении метаданных
	def onExtractMeta(self, path):
		global fileExt
		global metaKey
		global list_meta
		global fileType
		global metaType
		self.meta_data.clear()
		if (metaKey == False):
			list_meta = []
			self.combo_meta.clear()
			allKey = False
			# Извлечение метаданных изображения
			if (fileType == "image"):
				img = ImageMeta()
				list_img = img.extract_base(path)
				# Извлечение метаданных File
				if(len(list_img) != 0):
					if (allKey == False):
						list_meta.append('All')
						allKey = True
					list_meta.append('File')
					self.meta_data.addItem("# File:")
					i = 0
					for i in range(len(list_img)):
						self.meta_data.addItem(list_img[i])
				list_exif = img.extract_exif(path, metaType)
				zeroKey = False
				exifKey = False
				oneKey = False
				if (bool(list_exif) == True):
					if (allKey == False):
						list_meta.append('All')
						allKey = True
					for k, v in list_exif.items():
						# Извлечение метаданных 0th
						if(k[:4] == "0th:"):
							if (zeroKey == False):
								list_meta.append('0th')
								self.meta_data.addItem("# 0th:")
								zeroKey = True
							self.meta_data.addItem(k[4:] + ": " + str(v))
					for k, v in list_exif.items():
						# Извлечение метаданных Exif
						if(k[:5] == "Exif:"):
							if (exifKey == False):
								list_meta.append('Exif')
								self.meta_data.addItem("# Exif:")
								exifKey = True
							self.meta_data.addItem(k[5:] + ": " + str(v))
					for k, v in list_exif.items():
						# Извлечение метаданных 1th
						if(k[:4] == "1th:"):
							if (oneKey == False):
								list_meta.append('1th')
								self.meta_data.addItem("# 1th:")
								oneKey = True
							if (k[4:] == "xifImageWidth" or k[4:] == "xifImageLength"):
								self.meta_data.addItem("E" + k[4:] + ": " + str(v))
							else:
								self.meta_data.addItem(k[4:] + ": " + str(v))
				if (self.meta_data.count() != 0):
					self.combo_attr.setReadOnly(False)
				self.onInsertMeta()
			# Извлечение метаданных аудиофайлов
			elif (fileType == "audio"):
				audio = AudioMeta()
				list_audio = audio.extract_base(path)
				# Извлечение метаданных File
				if(len(list_audio) != 0):
					if (allKey == False):
						list_meta.append('All')
						allKey = True
					list_meta.append('File')
					self.meta_data.addItem("# File:")
					i = 0
					for i in range(len(list_audio)):
						self.meta_data.addItem(list_audio[i])
				list_id3 = audio.extract_id3(path, metaType)
				mpegKey = False
				id3Key = False
				if (bool(list_id3) == True):
					if (allKey == False):
						list_meta.append('All')
						allKey = True
					for k, v in list_id3.items():
						# Извлечение метаданных MPEG
						if(k[:5] == "MPEG:"):
							if (mpegKey == False):
								list_meta.append('MPEG')
								self.meta_data.addItem("# MPEG:")
								mpegKey = True
							self.meta_data.addItem(k[5:] + ": " + str(v))
					for k, v in list_id3.items():
						# Извлечение метаданных ID3
						if("ID3" in k[:3]):
							if (id3Key == False):
								#i = k.find(':')
								list_meta.append(k[:3])
								self.meta_data.addItem("# ID3:")
								id3Key = True
							self.meta_data.addItem(k[4:] + ": " + str(v))
				if (self.meta_data.count() != 0):
					self.combo_attr.setReadOnly(False)
				self.onInsertMeta()
			elif (fileType == "video"):
				video = VideoMeta()
				list_video = video.extract_base(path)
				# Извлечение метаданных File
				if(len(list_video) != 0):
					if (allKey == False):
						list_meta.append('All')
						allKey = True
					list_meta.append('File')
					self.meta_data.addItem("# File:")
					i = 0
					for i in range(len(list_video)):
						if (i != "Common:"):
							if(list_video[i] == "Video stream:"):
								list_meta.append('Video stream')
								self.meta_data.addItem("#" + list_video[i])
							elif(list_video[i] == "Audio stream:"):
								list_meta.append('Audio stream')
								self.meta_data.addItem("#" + list_video[i])
							else:
								self.meta_data.addItem(list_video[i])
				if (self.meta_data.count() != 0):
					self.combo_attr.setReadOnly(False)
				self.onInsertMeta()
			elif (fileType == "pdf"):
				pdf = PDFMeta()
				list_pdf = pdf.extract_pdf(path)
				# Извлечение метаданных PDF
				if (bool(list_pdf) == True):
					if (allKey == False):
						list_meta.append('All')
						allKey = True
					list_meta.append('PDF')
					self.meta_data.addItem("# PDF:")
					for k, v in list_pdf.items():
						self.meta_data.addItem(k + ": " + str(v))
				list_xmp = pdf.extract_xmp(path, metaType)
				pdfKey = False
				dcKey = False
				xmpKey = False
				xmpmmKey = False
				if (bool(list_xmp) == True):
					if (allKey == False):
						list_meta.append('All')
						allKey = True
					for k, v in list_xmp.items():
						# Извлечение метаданных XMP-pdf
						if(k[:8] == "XMP-pdf:"):
							if (pdfKey == False):
								list_meta.append('XMP-pdf')
								self.meta_data.addItem("# XMP-pdf:")
								pdfKey = True
							self.meta_data.addItem(k[8:] + ": " + str(v))
					for k, v in list_xmp.items():
						# Извлечение метаданных XMP-dc
						if(k[:7] == "XMP-dc:"):
							if (dcKey == False):
								list_meta.append('XMP-dc')
								self.meta_data.addItem("# XMP-dc:")
								dcKey = True
							self.meta_data.addItem(k[7:] + ": " + str(v))
					for k, v in list_xmp.items():
						# Извлечение метаданных XMP-xmp
						if(k[:8] == "XMP-xmp:"):
							if (xmpKey == False):
								list_meta.append('XMP-xmp')
								self.meta_data.addItem("# XMP-xmp:")
								xmpKey = True
							self.meta_data.addItem(k[8:] + ": " + str(v))
					for k, v in list_xmp.items():
						# Извлечение метаданных XMP-xmp
						if(k[:10] == "XMP-xmpmm:"):
							if (xmpmmKey == False):
								list_meta.append('XMP-xmpmm')
								self.meta_data.addItem("# XMP-xmpmm:")
								xmpmmKey = True
							self.meta_data.addItem(k[10:] + ": " + str(v))
				if (self.meta_data.count() != 0):
					self.combo_attr.setReadOnly(False)
				self.onInsertMeta()
			elif (fileType == "text files"):
				txt = TextMeta()
				list_txt = txt.extract_base(path)
				# Извлечение метаданных File
				if(len(list_txt) != 0):
					if (allKey == False):
						list_meta.append('All')
						allKey = True
					list_meta.append('File')
					self.meta_data.addItem("# File:")
					i = 0
					for i in range(len(list_txt)):
						self.meta_data.addItem(list_txt[i])
				if (self.meta_data.count() != 0):
					self.combo_attr.setReadOnly(False)
				self.onInsertMeta()
			elif (fileType == "archives"):
				arch = ArchivesMeta()
				list_arch = arch.extract_base(path)
				# Извлечение метаданных File
				if(len(list_arch) != 0):
					if (allKey == False):
						list_meta.append('All')
						allKey = True
					list_meta.append('File')
					self.meta_data.addItem("# File:")
					i = 0
					for i in range(len(list_arch)):
						self.meta_data.addItem(list_arch[i])
				list_zip = arch.extract_zip(path)
				# Извлечение метаданных File
				if(len(list_zip) != 0):
					if (allKey == False):
						list_meta.append('All')
						allKey = True
					list_meta.append('ZIP')
					self.meta_data.addItem("# ZIP:")
					i = 0
					for i in range(len(list_zip)):
						self.meta_data.addItem(list_zip[i])
				if (self.meta_data.count() != 0):
					self.combo_attr.setReadOnly(False)
				self.onInsertMeta()
			elif (fileType == "office documents"):
				off = OfficeMeta()
				list_off = off.extract_zip(path)
				# Извлечение метаданных File
				if(len(list_off) != 0):
					if (allKey == False):
						list_meta.append('All')
						allKey = True
					list_meta.append('ZIP')
					self.meta_data.addItem("# ZIP:")
					i = 0
					for i in range(len(list_off)):
						self.meta_data.addItem(list_off[i])
				list_d = off.extract_xml(path, fileExt ,metaType)
				dcKey = False
				xmlKey = False
				if (bool(list_d) == True):
					if (allKey == False):
						list_meta.append('All')
						allKey = True
					for k, v in list_d.items():
						if(k[:7] == "XML-dc:"):
							if (dcKey == False):
								list_meta.append('XML-dc')
								self.meta_data.addItem("# XML-dc:")
								dcKey = True
							self.meta_data.addItem(k[7:] + ": " + str(v))
					for k, v in list_d.items():
						if(k[:4] == "XML:"):
							if (xmlKey == False):
								list_meta.append('XML')
								self.meta_data.addItem("# XML:")
								xmlKey = True
							self.meta_data.addItem(k[4:] + ": " + str(v))
				if (self.meta_data.count() != 0):
					self.combo_attr.setReadOnly(False)
				self.onInsertMeta()
			elif (fileType == "executive files"):
				exe = EXEMeta()
				list_exe = exe.extract_base(path)
				if(len(list_exe) != 0):
					if (allKey == False):
						list_meta.append('All')
						allKey = True
					list_meta.append('File')
					self.meta_data.addItem("# File:")
					i = 0
					for i in range(len(list_exe)):
						self.meta_data.addItem(list_exe[i])
				if (self.meta_data.count() != 0):
					self.combo_attr.setReadOnly(False)
				self.onInsertMeta()
		elif (metaKey == True):
			if (fileType == "image"):
				img = ImageMeta()
				list_img = img.extract_base(path)
				if(metaType == "All" or metaType == "File"):
					if (len(list_img) != 0):
						self.meta_data.addItem("# File:")
						i = 0
						for i in range(len(list_img)):
							self.meta_data.addItem(list_img[i])
				list_exif = img.extract_exif(path, metaType)
				zeroKey = False
				exifKey = False
				oneKey = False
				if (bool(list_exif) == True):
					for k, v in list_exif.items():
						if (metaType == "All" or metaType == "0th"):
							if(k[:4] == "0th:"):
								if (zeroKey == False):
									self.meta_data.addItem("# 0th:")
									zeroKey = True
								self.meta_data.addItem(k[4:] + ": " + str(v))
					for k, v in list_exif.items():
						if (metaType == "All" or metaType == "Exif"):
							if(k[:5] == "Exif:"):
								if (exifKey == False):
									self.meta_data.addItem("# Exif:")
									exifKey = True
								self.meta_data.addItem(k[5:] + ": " + str(v))	
					for k, v in list_exif.items():
						if (metaType == "All" or metaType == "1th"):
							if(k[:4] == "1th:"):
								if (oneKey == False):
									self.meta_data.addItem("# 1th:")
									oneKey = True
								self.meta_data.addItem(k[4:] + ": " + str(v))
			elif (fileType == "audio"):
				audio = AudioMeta()
				list_audio = audio.extract_base(path)
				if(metaType == "All" or metaType == "File"):
					if (len(list_audio) != 0):
						self.meta_data.addItem("# File:")
						i = 0
						for i in range(len(list_audio)):
							self.meta_data.addItem(list_audio[i])
				list_id3 = audio.extract_id3(path, metaType)
				mpegKey = False
				id3Key = False
				if (bool(list_id3) == True):
					for k, v in list_id3.items():
						if (metaType == "All" or metaType == "MPEG"):
							if(k[:5] == "MPEG:"):
								if (mpegKey == False):
									self.meta_data.addItem("# MPEG:")
									mpegKey = True
								self.meta_data.addItem(k[5:] + ": " + str(v))
					for k, v in list_id3.items():
						if (metaType == "All" or "ID3" in metaType):
							if("ID3" in k[:3]):
								if (id3Key == False):
									#i = k.find(':')
									self.meta_data.addItem("# ID3:")
									id3Key = True
								self.meta_data.addItem(k[4:] + ": " + str(v))	
			elif (fileType == "video"):
				video = VideoMeta()
				list_video = video.extract_base(path)
				if(metaType == "All" or metaType == "File"):
					if (len(list_video) != 0):
						self.meta_data.addItem("# File:")
						i = 0
						for i in range(len(list_video)):
							if (list_video[i] != "Common:"):
								if(list_video[i] == "Video stream:"):
									break
								else:
									self.meta_data.addItem(list_video[i])
				if(metaType == "All" or metaType == "Video stream"):
					if (len(list_video) != 0):
						self.meta_data.addItem('# Video stream:')
						i = 0
						videoKey = False
						for i in range(len(list_video)):
							if (list_video[i] == "Video stream:"):
								videoKey = True
							elif (videoKey == True):
								if(list_video[i] == "Audio stream:"):
									break
								else:
									self.meta_data.addItem(list_video[i])
				if(metaType == "All" or metaType == "Audio stream"):
					if (len(list_video) != 0):
						self.meta_data.addItem('# Audio stream:')
						i = 0
						audioKey = False
						for i in range(len(list_video)):
							if (list_video[i] == "Audio stream:"):
								audioKey = True
							elif (audioKey == True):
								self.meta_data.addItem(list_video[i])
			elif (fileType == "pdf"):
				pdf = PDFMeta()
				list_pdf = pdf.extract_pdf(path)
				if(metaType == "All" or metaType == "PDF"):
					# Извлечение метаданных PDF
					if (bool(list_pdf) == True):
						self.meta_data.addItem("# PDF:")
						for k, v in list_pdf.items():
							self.meta_data.addItem(k + ": " + str(v))
				strok = metaType[4:]
				list_xmp = pdf.extract_xmp(path, strok)
				pdfKey = False
				dcKey = False
				xmpKey = False
				xmpmmKey = False
				print(bool(list_xmp))
				if (bool(list_xmp) == True):
					for k, v in list_xmp.items():
						if(metaType == "All" or metaType == "XMP-pdf"):
							# Извлечение метаданных XMP-pdf
							if(k[:8] == "XMP-pdf:"):
								if (pdfKey == False):
									self.meta_data.addItem("# XMP-pdf:")
									pdfKey = True
								self.meta_data.addItem(k[8:] + ": " + str(v))
					for k, v in list_xmp.items():
						if(metaType == "All" or metaType == "XMP-dc"):
							# Извлечение метаданных XMP-dc
							if(k[:7] == "XMP-dc:"):
								if (dcKey == False):
									self.meta_data.addItem("# XMP-dc:")
									dcKey = True
								self.meta_data.addItem(k[7:] + ": " + str(v))
					for k, v in list_xmp.items():
						if(metaType == "All" or metaType == "XMP-xmp"):
							# Извлечение метаданных XMP-xmp
							if(k[:8] == "XMP-xmp:"):
								if (xmpKey == False):
									self.meta_data.addItem("# XMP-xmp:")
									xmpKey = True
								self.meta_data.addItem(k[8:] + ": " + str(v))
					for k, v in list_xmp.items():
						if(metaType == "All" or metaType == "XMP-xmpmm"):
							# Извлечение метаданных XMP-xmp
							if(k[:10] == "XMP-xmpmm:"):
								if (xmpmmKey == False):
									self.meta_data.addItem("# XMP-xmpmm:")
									xmpmmKey = True
								self.meta_data.addItem(k[10:] + ": " + str(v))
			elif (fileType == "text files"):
				txt = TextMeta()
				list_txt = txt.extract_base(path)
				# Извлечение метаданных File
				if(len(list_txt) != 0):
					self.meta_data.addItem("# File:")
					i = 0
					for i in range(len(list_txt)):
						self.meta_data.addItem(list_txt[i])
			elif (fileType == "archives"):
				arch = ArchivesMeta()
				list_arch = arch.extract_base(path)
				if(metaType == "All" or metaType == "File"):
					# Извлечение метаданных File
					if(len(list_arch) != 0):
						self.meta_data.addItem("# File:")
						i = 0
						for i in range(len(list_arch)):
							self.meta_data.addItem(list_arch[i])
				# Извлечение метаданных File
				list_zip = arch.extract_zip(path)
				if(metaType == "All" or metaType == "ZIP"):
					if(len(list_zip) != 0):
						self.meta_data.addItem("# ZIP:")
						i = 0
						for i in range(len(list_zip)):
							self.meta_data.addItem(list_zip[i])
			elif (fileType == "office documents"):
				off = OfficeMeta()
				list_off = off.extract_zip(path)
				if(metaType == "All" or metaType == "ZIP"):
					if(len(list_off) != 0):
						self.meta_data.addItem("# ZIP:")
						i = 0
						for i in range(len(list_off)):
							self.meta_data.addItem(list_off[i])
				list_d = off.extract_xml(path, fileExt ,metaType)
				dcKey = False
				xmlKey = False
				if (bool(list_d) == True):
					for k, v in list_d.items():
						if(metaType == "All" or metaType == "XML-dc"):
							if(k[:7] == "XML-dc:"):
								if (dcKey == False):
									self.meta_data.addItem("# XML-dc:")
									dcKey = True
								self.meta_data.addItem(k[7:] + ": " + str(v))
					for k, v in list_d.items():
						if(metaType == "All" or metaType == "XML"):
							if(k[:4] == "XML:"):
								if (xmlKey == False):
									self.meta_data.addItem("# XML:")
									xmlKey = True
								self.meta_data.addItem(k[4:] + ": " + str(v))
			elif (fileType == "executive files"):
				exe = EXEMeta()
				list_exe = exe.extract_base(path)
				# Извлечение метаданных File
				if(len(list_exe) != 0):
					self.meta_data.addItem("# File:")
					i = 0
					for i in range(len(list_exe)):
						self.meta_data.addItem(list_exe[i])
			metaKey = False
		if(self.meta_data.count() != 0):
			self.combo_attr.setReadOnly(False)
		else:
			self.combo_attr.setReadOnly(True)

	# Функция изменения метаданных
	def onUpdateAttr(self, *args):
		global filePath
		global fileName
		global fileExt
		global fileType
		if (self.combo_attr.toPlainText() == ""):
			btnReply = QMessageBox.warning(self, 'Предупреждение', "Введите значение для изменения!", QMessageBox.Ok)
		else:
			if (fileType == "image"):
				btnReply = QMessageBox.question(self, 'Вопрос', "Вы действительно хотите изменить элемент?", QMessageBox.Yes | QMessageBox.No)
				if (btnReply == QMessageBox.Yes):
					it = ImageMeta()
					if (self.combo_attr.toPlainText() == "xifImageWidth" or self.combo_attr.toPlainText() == "xifImageLength"):
						result = it.update_exif(filePath + '/' + fileName + fileExt, "E" + self.combo_attr.toPlainText(), self.new_value.toPlainText())
					else:
						result = it.update_exif(filePath + '/' + fileName + fileExt, self.combo_attr.toPlainText(), self.new_value.toPlainText())
					if (result == 1):
						self.onExtractMeta(filePath + '/' + fileName + fileExt)
						btnReply = QMessageBox.information(self, 'Результат', "Элемент успешно изменен!", QMessageBox.Ok)
					elif (result == 0):
						btnReply = QMessageBox.warning(self, 'Предупреждение', "Данный элемент изменить невозможно!", QMessageBox.Ok)
					elif (result == -1):
						btnReply = QMessageBox.critical(self, 'Ошибка', "Неправильно введено значение элемента!", QMessageBox.Ok)
			if (fileType == "audio"):
				btnReply = QMessageBox.question(self, 'Вопрос', "Вы действительно хотите изменить элемент?", QMessageBox.Yes | QMessageBox.No)
				if (btnReply == QMessageBox.Yes):
					it = AudioMeta()
					result = it.update_id3(filePath + '/' + fileName + fileExt, self.combo_attr.toPlainText(), self.new_value.toPlainText())
					if (result == 1):
						self.onExtractMeta(filePath + '/' + fileName + fileExt)
						btnReply = QMessageBox.information(self, 'Результат', "Элемент успешно изменен!", QMessageBox.Ok)
					elif (result == 0):
						btnReply = QMessageBox.warning(self, 'Предупреждение', "Данный элемент изменить невозможно!", QMessageBox.Ok)
					elif (result == -1):
						btnReply = QMessageBox.critical(self, 'Ошибка', "Неправильно введено значение элемента!", QMessageBox.Ok)
			if (fileType == "video"):
				btnReply = QMessageBox.question(self, 'Вопрос', "Вы действительно хотите изменить элемент?", QMessageBox.Yes | QMessageBox.No)
				if (btnReply == QMessageBox.Yes):
					btnReply = QMessageBox.warning(self, 'Предупреждение', "Данный элемент изменить невозможно!", QMessageBox.Ok)
			if (fileType == "pdf"):
				btnReply = QMessageBox.question(self, 'Вопрос', "Вы действительно хотите изменить элемент?", QMessageBox.Yes | QMessageBox.No)
				if (btnReply == QMessageBox.Yes):
					it = PDFMeta()
					result = it.update_pdf(filePath + '/' + fileName + fileExt, self.combo_attr.toPlainText(), self.new_value.toPlainText())
					if (result == 1):
						self.onExtractMeta(filePath + '/' + fileName + fileExt)
						btnReply = QMessageBox.information(self, 'Результат', "Элемент успешно изменен!", QMessageBox.Ok)
					elif (result == 0):
						btnReply = QMessageBox.warning(self, 'Предупреждение', "Данный элемент изменить невозможно!", QMessageBox.Ok)
					elif (result == -1):
						btnReply = QMessageBox.critical(self, 'Ошибка', "Неправильно введено значение элемента!", QMessageBox.Ok)
			if (fileType == "text files"):
				btnReply = QMessageBox.question(self, 'Вопрос', "Вы действительно хотите изменить элемент?", QMessageBox.Yes | QMessageBox.No)
				if (btnReply == QMessageBox.Yes):
					btnReply = QMessageBox.warning(self, 'Предупреждение', "Данный элемент изменить невозможно!", QMessageBox.Ok)
			if (fileType == "archives"):
				btnReply = QMessageBox.question(self, 'Вопрос', "Вы действительно хотите изменить элемент?", QMessageBox.Yes | QMessageBox.No)
				if (btnReply == QMessageBox.Yes):
					btnReply = QMessageBox.warning(self, 'Предупреждение', "Данный элемент изменить невозможно!", QMessageBox.Ok)
			if (fileType == "office documents"):
				btnReply = QMessageBox.question(self, 'Вопрос', "Вы действительно хотите изменить элемент?", QMessageBox.Yes | QMessageBox.No)
				if (btnReply == QMessageBox.Yes):
					it = OfficeMeta()
					if (fileExt == ".doc" or fileExt == ".docx"):
						result = it.update_docx(filePath + '/' + fileName + fileExt, self.combo_attr.toPlainText(), self.new_value.toPlainText())
					elif (fileExt == ".ppt" or fileExt == ".pptx"):
						result = it.update_pptx(filePath + '/' + fileName + fileExt, self.combo_attr.toPlainText(), self.new_value.toPlainText())
					elif (fileExt == ".xls" or fileExt == ".xlsx"):
						result = it.update_xlsx(filePath + '/' + fileName + fileExt, self.combo_attr.toPlainText(), self.new_value.toPlainText())
					if (result == 1):
						self.onExtractMeta(filePath + '/' + fileName + fileExt)
						btnReply = QMessageBox.information(self, 'Результат', "Элемент успешно изменен!", QMessageBox.Ok)
					elif (result == 0):
						btnReply = QMessageBox.warning(self, 'Предупреждение', "Данный элемент изменить невозможно!", QMessageBox.Ok)
					elif (result == -1):
						btnReply = QMessageBox.critical(self, 'Ошибка', "Неправильно введено значение элемента!", QMessageBox.Ok)
			if (fileType == "executive files"):
				btnReply = QMessageBox.question(self, 'Вопрос', "Вы действительно хотите изменить элемент?", QMessageBox.Yes | QMessageBox.No)
				if (btnReply == QMessageBox.Yes):
					btnReply = QMessageBox.warning(self, 'Предупреждение', "Данный элемент изменить невозможно!", QMessageBox.Ok)

	# Функция удаления метаданных
	def onDelAttr(self, *args):
		global filePath
		global fileName
		global fileExt
		global fileType
		if (self.combo_attr.toPlainText() == ""):
			btnReply = QMessageBox.warning(self, 'Предупреждение', "Выберите элемент для удаления!", QMessageBox.Ok)
		else:
			if (fileType == "image"):
				btnReply = QMessageBox.question(self, 'Вопрос', "Вы действительно хотите удалить элемент?", QMessageBox.Yes | QMessageBox.No)
				if (btnReply == QMessageBox.Yes):
					btnReply = QMessageBox.warning(self, 'Предупреждение', "Данный элемент удалить невозможно!", QMessageBox.Ok)
			if (fileType == "audio"):
				btnReply = QMessageBox.question(self, 'Вопрос', "Вы действительно хотите удалить элемент?", QMessageBox.Yes | QMessageBox.No)
				if (btnReply == QMessageBox.Yes):
					btnReply = QMessageBox.warning(self, 'Предупреждение', "Данный элемент удалить невозможно!", QMessageBox.Ok)
			if (fileType == "video"):
				btnReply = QMessageBox.question(self, 'Вопрос', "Вы действительно хотите удалить элемент?", QMessageBox.Yes | QMessageBox.No)
				if (btnReply == QMessageBox.Yes):
					btnReply = QMessageBox.warning(self, 'Предупреждение', "Данный элемент удалить невозможно!", QMessageBox.Ok)
			if (fileType == "pdf"):
				btnReply = QMessageBox.question(self, 'Вопрос', "Удаление элемента повлечет удаление XMP\nВы действительно хотите удалить элемент?", QMessageBox.Yes | QMessageBox.No)
				if (btnReply == QMessageBox.Yes):
					it = PDFMeta()
					result = it.remove_pdf(filePath + '/' + fileName + fileExt, self.combo_attr.toPlainText())
					self.onExtractMeta(filePath + '/' + fileName + fileExt)
					if (result == 1):
						btnReply = QMessageBox.information(self, 'Результат', "Элемент успешно удален!", QMessageBox.Ok)
					else:
						btnReply = QMessageBox.warning(self, 'Предупреждение', "Данный элемент удалить невозможно!", QMessageBox.Ok)
			if (fileType == "text files"):
				btnReply = QMessageBox.question(self, 'Вопрос', "Вы действительно хотите удалить элемент?", QMessageBox.Yes | QMessageBox.No)
				if (btnReply == QMessageBox.Yes):
					btnReply = QMessageBox.warning(self, 'Предупреждение', "Данный элемент удалить невозможно!", QMessageBox.Ok)
			if (fileType == "archives"):
				btnReply = QMessageBox.question(self, 'Вопрос', "Вы действительно хотите удалить элемент?", QMessageBox.Yes | QMessageBox.No)
				if (btnReply == QMessageBox.Yes):
					btnReply = QMessageBox.warning(self, 'Предупреждение', "Данный элемент удалить невозможно!", QMessageBox.Ok)
			if (fileType == "office documents"):
				btnReply = QMessageBox.question(self, 'Вопрос', "Вы действительно хотите удалить элемент?", QMessageBox.Yes | QMessageBox.No)
				if (btnReply == QMessageBox.Yes):
					btnReply = QMessageBox.warning(self, 'Предупреждение', "Данный элемент удалить невозможно!", QMessageBox.Ok)
			if (fileType == "executive files"):
				btnReply = QMessageBox.question(self, 'Вопрос', "Вы действительно хотите удалить элемент?", QMessageBox.Yes | QMessageBox.No)
				if (btnReply == QMessageBox.Yes):
					btnReply = QMessageBox.warning(self, 'Предупреждение', "Данный элемент удалить невозможно!", QMessageBox.Ok)
	
	# Функция удаления группы метаданных
	def onDelMeta(self, *args):
		global filePath
		global fileName
		global fileExt
		global fileType
		global metaType
		if (self.combo_meta.currentText() == "" or self.meta_data.count() == 0):
			btnReply = QMessageBox.warning(self, 'Предупреждение', "Выведите метаданные для удаления!", QMessageBox.Ok)
		else:
			if (fileType == "image"):
				btnReply = QMessageBox.question(self, 'Вопрос', "Вы действительно хотите удалить метаданные?", QMessageBox.Yes | QMessageBox.No)
				if (btnReply == QMessageBox.Yes):
					if (metaType == "0th" or metaType == "Exif" or metaType == "1th"):
						it = ImageMeta()
						it.remove_exif_all(filePath + '/' + fileName + fileExt)
						metaType = ""
						self.onExtractMeta(filePath + '/' + fileName + fileExt)
						btnReply = QMessageBox.information(self, 'Результат', "Метаданные успешно удалены!", QMessageBox.Ok)
					else:
						btnReply = QMessageBox.warning(self, 'Предупреждение', "Данные метаданные удалить невозможно!", QMessageBox.Ok)
			if (fileType == "audio"):
				btnReply = QMessageBox.question(self, 'Вопрос', "Вы действительно хотите удалить метаданные?", QMessageBox.Yes | QMessageBox.No)
				if (btnReply == QMessageBox.Yes):
					if ("ID3" in metaType):
						it = AudioMeta()
						it.remove_id3_all(filePath + '/' + fileName + fileExt)
						metaType = ""
						self.onExtractMeta(filePath + '/' + fileName + fileExt)
						btnReply = QMessageBox.information(self, 'Результат', "Метаданные успешно удалены!", QMessageBox.Ok)
					else:
						btnReply = QMessageBox.warning(self, 'Предупреждение', "Данные метаданные удалить невозможно!", QMessageBox.Ok)
			if (fileType == "video"):
				btnReply = QMessageBox.question(self, 'Вопрос', "Вы действительно хотите удалить метаданные?", QMessageBox.Yes | QMessageBox.No)
				if (btnReply == QMessageBox.Yes):
					btnReply = QMessageBox.warning(self, 'Предупреждение', "Данные метаданные удалить невозможно!", QMessageBox.Ok)
			if (fileType == "pdf"):
				btnReply = QMessageBox.question(self, 'Вопрос', "Удаление каснется всех метаданных\nВы действительно хотите удалить метаданные?", QMessageBox.Yes | QMessageBox.No)
				if (btnReply == QMessageBox.Yes):
					it = PDFMeta()
					it.remove_pdf_all(filePath + '/' + fileName + fileExt)
					metaType = ""
					self.onExtractMeta(filePath + '/' + fileName + fileExt)
					btnReply = QMessageBox.information(self, 'Результат', "Метаданные успешно удалены!", QMessageBox.Ok)
			if (fileType == "text files"):
				btnReply = QMessageBox.question(self, 'Вопрос', "Вы действительно хотите удалить метаданные?", QMessageBox.Yes | QMessageBox.No)
				if (btnReply == QMessageBox.Yes):
					btnReply = QMessageBox.warning(self, 'Предупреждение', "Данные метаданные удалить невозможно!", QMessageBox.Ok)
			if (fileType == "archives"):
				btnReply = QMessageBox.question(self, 'Вопрос', "Вы действительно хотите удалить метаданные?", QMessageBox.Yes | QMessageBox.No)
				if (btnReply == QMessageBox.Yes):
					btnReply = QMessageBox.warning(self, 'Предупреждение', "Данные метаданные удалить невозможно!", QMessageBox.Ok)
			if (fileType == "office documents"):
				btnReply = QMessageBox.question(self, 'Вопрос', "Вы действительно хотите удалить метаданные?", QMessageBox.Yes | QMessageBox.No)
				if (btnReply == QMessageBox.Yes):
					btnReply = QMessageBox.warning(self, 'Предупреждение', "Данные метаданные удалить невозможно!", QMessageBox.Ok)
			if (fileType == "executive files"):
				btnReply = QMessageBox.question(self, 'Вопрос', "Вы действительно хотите удалить метаданные?", QMessageBox.Yes | QMessageBox.No)
				if (btnReply == QMessageBox.Yes):
					btnReply = QMessageBox.warning(self, 'Предупреждение', "Данные метаданные удалить невозможно!", QMessageBox.Ok)
					

	# Функция сохранения лога метаданных
	def onSaveMeta(self, *args):
		global filePath
		global fileName
		if (self.combo_meta.currentText() == "" or self.meta_data.count() == 0):
			btnReply = QMessageBox.warning(self, 'Предупреждение', "Выведите метаданные для сохранения!", QMessageBox.Ok)
		else:
			btnReply = QMessageBox.question(self, 'Вопрос', "Вы действительно хотите сохранить метаданные?", QMessageBox.Yes | QMessageBox.No)
			if (btnReply == QMessageBox.Yes):
				with open(filePath + "/" + fileName + ".txt", "w") as file:
					for i in range(self.meta_data.count()):
						file.write(str(self.meta_data.item(i).text()) + "\n")
					file.close()
				btnReply = QMessageBox.information(self, 'Результат', "Метаданные успешно сохранены!", QMessageBox.Ok)

if __name__ == '__main__':
	app = QApplication(sys.argv)
	dirPath = r'<Your directory>'
	filePath = "" # Путь до файла
	fileName = "" # Имя файла
	fileFormat = ["*.*"] # Список расширений типа файла
	fileExt = ".*" # Расширение файла
	fileType = "" # Тип файла
	list_meta = [] # Список типов метаданных
	metaKey = False # Ключ при выборе типа метаданных
	metaType = "" # Тип метаданных
	
	delKey = False
	
	metaSoft = MetaDataView(dirPath)
	metaSoft.show()
	sys.exit(app.exec_())