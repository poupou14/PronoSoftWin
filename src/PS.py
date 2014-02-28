#!/usr/bin/python 
import string, sys
from PSParser import PSParser
#from PSWriter import PSWriter

def main():
	if len(sys.argv) == 2 :
		if sys.argv[1] == "-h" :
			print "user help :"
			print "$ python ./PS.py <grille> <nbMatches> <annee> <from> <to>"
			print "\t grille : numero de la grille"
			print "\t nbMatches : 7 ou 15"
			print "\t annee : annnee de la grille" 
			print "\t from : a partir du joueur numero from" 
			print "\t to : jusqu au joueur numero to" 
	elif len(sys.argv) == 6 :
		myPS = PSParser()
#	myPSWriter = PSWriter("../INPUT/PS_Players.xls")
		#sourceFile_l = "/home/lili/Developpement/HockeyLinks/INPUT/InpotPS.xls"
		#targetFile_l = "/home/lili/Developpement/HockeyLinks/OUTPUT/PronoSoft.xls"
		grille_l = sys.argv[1]	
		nbGames_l = sys.argv[2]	
		annee_l = sys.argv[3]	
		from_l = sys.argv[4]	
		to_l = sys.argv[5]	
		myPS.readPS(grille_l, int(nbGames_l), annee_l, int(from_l), int(to_l))

main()
