#!/usr/bin/python 
from HTMLParser import HTMLParser
import PSDataFormat
import os,string, sys
import urllib
import urllib2
import copy
#import chardet
#### SPECIFIC IMPORT #####
sys.path.append("../Import/xlrd-0.7.1")
sys.path.append("../Import/xlwt-0.7.2")
sys.path.append("../Import/pyexcelerator-0.6.4.1")

from pyExcelerator import *
import xlwt
from xlrd import open_workbook
from xlwt import Workbook,easyxf,Formula,Style
#from lxml import etree
import xlrd

currentUser = dict()
grille = PSDataFormat.emptyGrille
grilleEmpty = True


def onlyascii(char):
    if ord(char) <= 0 or ord(char) > 127: 
	return ''
    else: 
	return char

def isnumber(s):
    try:
        float(s)
        return True
    except ValueError:
        return False

class PSParser():

	def __init__(self): 
		self.__status = False
		self.__userCounter = 0
		self.__fileUserCounter = 0
		self.__userNext = True
		self.__pronoLink = "http://www.pronosoft.com/"
		self.__prePage7 = "http://www.pronosoft.com/fr/concours/concours_lotofoot.php?lf7&mode=pronos&id7=1"
		self.__prePage = "http://www.pronosoft.com/fr/concours/concours_lotofoot.php?lf7&mode=pronos&id7=1"
		self.__prePage15 = "http://www.pronosoft.com/fr/concours/concours_lotofoot.php?&mode=pronos&id15=1"
		self.__midPage = "&sort=date&start="
		self.__sufPage = "#pronosc"
		self.__to = 0
		self.__nbGames = 0
		self.__from = 0
    		self.__workbook1 = Workbook()
    		self.__grilleSheet = self.__workbook1.add_sheet("Grille", cell_overwrite_ok=True)
		self.__outPutFileName = ""
		self.__grille = ""
		self.__annee = ""

	def findRootPage(self):
		notRead_l = True
		myPronoParser = PSPronoParser()
		myPronoParser.setGrille(self.__grille)
		myPronoParser.setAnnee(self.__annee)
		psRootUrl_l = ''.join((self.__prePage, self.__midPage))
		psRootUrl_l = ''.join((psRootUrl_l, "%d" % self.__userCounter))
		psRootUrl_l = ''.join((psRootUrl_l, self.__sufPage))
		myPronoParser.reset()
		while notRead_l :
			try :
				print "Ouverture : %s" % psRootUrl_l
				url = urllib2.urlopen(psRootUrl_l, timeout = 5)
				print "Lecture : %s" % psRootUrl_l
				myPronoParser.html = url.read()
				notRead_l = False
			#except IOError :
			except :
				notRead_l = True
				print "url read issue"
		url.close	
		print "Fermeture : %s" % psRootUrl_l
		myPronoParser.html = filter(onlyascii, myPronoParser.html)
		myPronoParser.feed(myPronoParser.html)	
		prePage_l = myPronoParser.getRootAddress()
		prePage_l = prePage_l.replace("#menuc", "")
		self.__prePage = ''.join((self.__pronoLink, prePage_l)) 
		
		print "root page : %s" % self.__prePage


	def readPS(self, grille_p, nbGames_p, annee_p, from_p, to_p):
		global currentUser
		# On lit la premiere page puis les suivantes jusqu'a la page vide
		notRead_l = True
		self.__userCounter = from_p
		self.__fileUserCounter = from_p
		self.__from = from_p
		self.__to = to_p
		self.__outPutFileName = "Grille_%d" % nbGames_p
		self.__outPutFileName = ''.join((self.__outPutFileName, "-%s.xls" % grille_p))
		
		self.__grille = grille_p	
		self.__annee = annee_p
		
		if nbGames_p == 15 :
			self.__prePage = self.__prePage15
			self.__nbGames = 14
		else :
			self.__prePage = self.__prePage7
			self.__nbGames = 7

		self.findRootPage()

		#while (False) :
		while (self.__userNext) :
			myUserParser = PSUserParser()
			psUserUrl_l = ''.join((self.__prePage, self.__midPage))
			psUserUrl_l = ''.join((psUserUrl_l, "%d" % self.__userCounter))
			psUserUrl_l = ''.join((psUserUrl_l, self.__sufPage))
			currentUser['prono'] = []
			myUserParser.reset()
			myUserParser.setNbGames(self.__nbGames)
			notRead_l = True
			try :
				while notRead_l :
					try :
						print "Ouverture : %s" % psUserUrl_l
						url = urllib2.urlopen(psUserUrl_l, timeout = 5)
						print "Lecture : %s" % psUserUrl_l
						myUserParser.html = url.read()
						notRead_l = False
					#except IOError :
					except :
						notRead_l = True
						print "url read issue"
				url.close()
				print "Fermeture : %s" % psUserUrl_l
				myUserParser.html = filter(onlyascii, myUserParser.html)
				myUserParser.feed(myUserParser.html)	
				PSDataFormat.listUser.append(currentUser.copy())
				self.__userCounter += 1
			except IOError:
				print "problem while reading %s" % psUserUrl_l
				self.__userNext = False
			#print PSDataFormat.listUser
			if (self.__to > self.__from) :
				self.__userNext = (self.__userCounter <= to_p)
			if self.__userCounter % 10 == 0 :
				print "user %d, save xcl file" % self.__userCounter
				self.writeOuput()	
			self.__userNext = self.__userNext and not myUserParser.getLastPage() 

		self.writeOuput()	




	def writeOuput(self):
#		Book to read xls file (output of main_CSVtoXLS)

		newUserIndex_l = 1
		title = False
			
		# affichage des titres si ce n'est deja fait
		
		if (self.__fileUserCounter == 0) :
			self.addGames()

		for userIndex_l in range(self.__fileUserCounter, self.__userCounter) :
			user_l = PSDataFormat.listUser[userIndex_l-self.__fileUserCounter]
			self.addUser(user_l, userIndex_l-self.__fileUserCounter) 
		self.__workbook1.save(self.__outPutFileName)

	def addGames(self) :
		indexClmn_l = 0
		self.__grilleSheet.write(0, 0, grille['Titre'])
		#for line_l in range(0, len(grille['match'])-1) :
		for clmn_l in range(0, len(grille['match'])) :
			value_l = filter(onlyascii, grille['match'][clmn_l])
			self.__grilleSheet.write(0, clmn_l + 2, value_l)
 	
	def addUser(self, player_p, index_p):
		indexClmn_l = 0
		value_l = filter(onlyascii, PSDataFormat.listUser[index_p]['Name'])
		self.__grilleSheet.write(self.__fileUserCounter + index_p + 1, 0, value_l)
		value2_l = filter(onlyascii, PSDataFormat.listUser[index_p]['Date'])
		self.__grilleSheet.write(self.__fileUserCounter + index_p + 1, 1, value2_l)
		#style = easyxf('font: underline single')
#		for key in player.keys() :
		for indexClmn_l in range(0, len(player_p['prono'])) :
		#for indexClmn_l in range(0, len(player_p['prono'])-1) :
			value_l = filter(onlyascii, player_p['prono'][indexClmn_l])
			self.__grilleSheet.write(self.__fileUserCounter + index_p + 1, indexClmn_l + 2, value_l)
				
			
class PSUserParser(HTMLParser): 

	def __init__(self): 
		self.__readOK = False 
		self.__getTitle = False 
		self.__betweenTag = False
		self.__title = False 
		self.__nextUser = False 
		self.__nextDate = 0
		self.__nextTitle = False
		self.__nextGame = [0]
		self.__gameId = 0
		self.__nbGames = 0
		self.__nbProno = 0
		self.__lastPage = False


	def setReadNext(self, val) :
		self.__readNext = val

	def setNbGames(self, val) :
		self.__nbGames = val

	def readNext(self) :
		return self.__readNext	

	def getLastPage(self) :
		return self.__lastPage	

	def handle_starttag(self, tag, attrs):
		global currentUser
		if tag == "th" and self.__nextTitle :
			self.__nextTitle = True
		elif tag == "h2" and len(attrs) == 2 :
			if attrs[0][0] == "id" : 
				self.__nextUser = True
		elif tag == "tr" and len(attrs) == 2 and self.__nextDate == 0 :
			if attrs[0][0] == "id" :
				self.__nextDate = 1
		elif tag == "td" and len(attrs) == 1 and self.__nextDate == 1 :
			if attrs[0][0] == "class" :
				self.__nextDate = 2
		elif tag == "td" and len(attrs) == 1 :
			if self.__nbProno == 3 :
				self.__gameId += 1
				self.__nbProno = 0
				self.__nextGame.append(1)
			if attrs[0][1] == "grille_av" and self.__nextGame[self.__gameId] == 1 :
				self.__nextGame[self.__gameId] = 2			
			if attrs[0][1] == "grille_av" and self.__nextGame[self.__gameId] == 3 :
				self.__nextGame[self.__gameId] = 4			
			#elif attrs[0][1] == "grille" and attrs[0][0] == "class" :
				#self.__gameId += 1
				#self.__nextGame.append(1)
				#currentUser['prono'].append("")
				#print self.__gameId
		elif tag == "input" and len(attrs) == 6 :
			if attrs[1][1].find("prono") != -1 :
				self.__nbProno += 1
				#print "self.__nbProno = %d" % self.__nbProno
			if attrs[1][1].find("prono") != -1 and attrs[1][1].rfind("_0") > 6 and attrs[0][1].find("checkbox") != -1 :
				#print attrs
				currentUser['prono'].append("1")
			elif attrs[1][1].find("prono") != -1 and attrs[1][1].rfind("_1") > 6 and attrs[0][1].find("checkbox") != -1 :
				#print attrs
				currentUser['prono'][self.__gameId] = ''.join((currentUser['prono'][self.__gameId], "N"))
			elif attrs[1][1].find("prono") != -1 and attrs[1][1].rfind("_2") > 6 and attrs[0][1].find("checkbox") != -1 :
				#print attrs
				currentUser['prono'][self.__gameId] = ''.join((currentUser['prono'][self.__gameId], "2"))
		elif tag == "input" and len(attrs) == 5 :
			if attrs[1][1].find("prono")  != -1:
				self.__nbProno += 1
				#print "self.__nbProno = %d" % self.__nbProno
			if attrs[1][1].rfind("_0", -2) > 6 and attrs[0][1].find("checkbox") != -1 :
				#print attrs
				currentUser['prono'].append("")
		elif tag == "input" and len(attrs) == 4 : # match gagnant
			if attrs[1][1].find("_9")  != -1:
				currentUser['prono'].append("GAGNANT")
				self.__nbProno = 3
			#print currentUser
		self.__betweenTag = True

	def handle_data(self, data):
		global currentUser
		global grilleEmpty
		global grille
		if data.find("Aucun pronostic disponible") != -1 :
			print "fin du site !"
			self.__lastPage = True
		elif self.__nextUser :
			currentUser['Name'] = data
			self.__nextUser = False
			self.__nextTitle = True
		elif self.__nextDate == 2 :
			print "data : %s" % data
			self.__nextDate = 0
			currentUser['Date'] = data
		elif self.__nextTitle :
			if grilleEmpty :
				grille['Titre'] = data
			self.__nextTitle = False
			self.__nextGame[0] = 1
		elif self.__nextGame[self.__gameId] == 2 :# equipe A du match __gameId
			if grilleEmpty :
				grille['match'].append(data)
				print grille['match'][self.__gameId]
			self.__nextGame[self.__gameId] = 3
		elif self.__nextGame[self.__gameId] == 4 :# equipe B du match __gameId
			if grilleEmpty :
				equipeA_l = grille['match'][self.__gameId]
				equipeA_l = ''.join((equipeA_l, '-'))
				grille['match'][self.__gameId] = ''.join((equipeA_l, data))
				print grille['match'][self.__gameId]
				if self.__gameId == self.__nbGames-1 :
					grilleEmpty = False
				#self.__gameId += 1
				#self.__nextGame.append(1)
				
			self.__nextGame[self.__gameId] = 5
			#print "next game : %s" % self.__nextGame


	def handle_endtag(self, tag):
		self.__betweenTag = False
		#if self.__nbProno == 3 :
			#self.__gameId += 1
			#self.__nbProno = 0
			#self.__nextGame.append(1)

class PSPronoParser(HTMLParser): 

	def __init__(self): 
		self.__betweenTag = False
		self.__title = False 
		self.__valueTmp = ""
		self.__value = ""
		self.__grille = ""
		self.__annee = ""

	def setAnnee(self, annee_p) :
		self.__annee = annee_p

	def setGrille(self, grille_p) :
		#self.__grille = ''.join((';',grille_p))
		self.__grille = grille_p

	def getRootAddress(self) :
		return self.__value

	def handle_starttag(self, tag, attrs):
		global currentUser
		if tag == "option" and len(attrs) == 1 :
			if attrs[0][0] == "value" : 
				self.__valueTmp = attrs[0][1]
		elif tag == "option" and len(attrs) == 2 :
			if attrs[1][0] == "value" : 
				self.__valueTmp = attrs[1][1]
		self.__betweenTag = True

	def handle_data(self, data):
		grilleSearch_l = "%s du" % self.__grille	
		if (data.find(grilleSearch_l) == 0) and (data.find(self.__annee) != -1) :
		#if data.find(self.__grille) != -1 :
			print "root adress : %s" % self.__valueTmp
			print "data : %s" % data
			print "index : %d" % data.find(self.__grille)
			self.__value = self.__valueTmp
		
			

	def handle_endtag(self, tag):
		self.__betweenTag = False





def open_excel_sheet():
    """ Opens a reference to an Excel WorkBook and Worksheet objects """
    workbook = Workbook()
    worksheet = workbook.add_sheet("Sheet 1")
    return workbook, worksheet

def write_excel_header(worksheet, title_cols):
    """ Write the header line into the worksheet """
    cno = 0
    for title_col in title_cols:
        worksheet.write(0, cno, title_col)
        cno = cno + 1
    return

def write_excel_row(worksheet, rowNumber, columnNumber):
    """ Write a non-header row into the worksheet """
    cno = 0
    for column in columns:
        worksheet.write(lno, cno, column)
        cno = cno + 1
    return

def save_excel_sheet(workbook, output_file_name):
    """ Saves the in-memory WorkBook object into the specified file """
    workbook.save(output_file_name)
    return

