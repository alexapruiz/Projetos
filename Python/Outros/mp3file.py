#!/usr/bin/python
# o Module name mp3file
#	o Classes : BigBooBoo : Class for exceptions/errors/boo-boo's
#	            mp3file   : Class representing an mp3file, is derived from
#	                        Userdict.Userdict
#========================================================================
# For the test, usage : mp3file [filenames]
# 	If no filenames are given, it assumes all files in the present directory.
# This module is part of an application I'd like to write that would let you
# read/write/modify the ID3 tags of mp3 files, it will also let you change the
# file name, and organize your mp3 files and all this would be done with a
# curses front-end. This module does all those things (except the "organize" part).
# If you'd like me explain anything or change anything or if you'd like to join me
# in finishing this thing up ....feel free to contact me: steve<lonetwin@yahoo.com>
#  *I've written+tested this on Linux only, though I see no reason why it
#   shouldn't work on windows.
# Note: 
# There is another nice 'lil script I've written called chgnm.py, it
# "...changes the names of files named the way I don't like 'em to the way I like
# 'em...", it can be found at http://www.lowerstandard.com/python/chgnm.py, This
# site by the way (The Useless Python site) is a *real* good site, lots of code
# there that is "...far from Useless...", if you are just learning python (like
# me :D )...you'd definitely want to have a look at it, the URL is the one above
# minus the 'chgnm.py'. 'nuff said, Have fun !!!
#==========================================================================

from UserDict import UserDict
import re
import os

def stripnulls(data):
	"strip whitespace and nulls"
	return data.replace("\0", " ").strip()

class BigBooBoo:
	def __init__(self, str='You crash application !! You Bad !!!', err=0):
		self.str = str
		self.err = err
	def __str__(self):
		return err, str
		
class mp3file(UserDict):
	tagDataMap = { "title"   : (  3,  33, stripnulls),
	               "artist"  : ( 33,  63, stripnulls),
	               "album"   : ( 63,  93, stripnulls),
	               "year"    : ( 93,  97, stripnulls),
	               "comment" : ( 96, 126, stripnulls),
	               "genre"   : (127, 128, ord) }

	rulez = [ ('One or more underscores',              '_+',       [ 1, ' ']),
	          ('Two or more hyphens',                   '-{2,}',    [ 1, '-']),
	          ('An ampersand',                         '&',        [ 1, 'And']),
	          ('Words that have a hyphen between them', '(-)(\w*)', [ 1, r' \1 \2'])	]

	# Extracted from mpg123 0.59q
	glist = [
	  "Blues", "Classic Rock", "Country", "Dance", "Disco", "Funk", "Grunge", "Hip-Hop",
	  "Jazz", "Metal", "New Age", "Oldies", "Other", "Pop", "R&B", "Rap", "Reggae", "Rock",
	  "Techno", "Industrial", "Alternative", "Ska", "Death Metal", "Pranks", "Soundtrack",
	  "Euro-Techno", "Ambient", "Trip-Hop", "Vocal", "Jazz+Funk", "Fusion", "Trance",
	  "Classical", "Instrumental", "Acid", "House", "Game", "Sound Clip", "Gospel", "Noise",
	  "Alt", "Bass", "Soul", "Punk", "Space", "Meditative", "Instrumental Pop",
	  "Instrumental Rock", "Ethnic", "Gothic", "Darkwave", "Techno-Industrial",
	  "Electronic", "Pop-Folk", "Eurodance", "Dream", "Southern Rock", "Comedy", "Cult",
	  "Gangsta Rap", "Top 40", "Christian Rap", "Pop/Funk", "Jungle", "Native American",
	  "Cabaret", "New Wave", "Psychedelic", "Rave", "Showtunes", "Trailer", "Lo-Fi",
	  "Tribal", "Acid Punk", "Acid Jazz", "Polka", "Retro", "Musical", "Rock & Roll",
	  "Hard Rock", "Folk", "Folk/Rock", "National Folk", "Swing", "Fast-Fusion", "Bebob",
	  "Latin", "Revival", "Celtic", "Bluegrass", "Avantgarde", "Gothic Rock",
	  "Progressive Rock", "Psychedelic Rock", "Symphonic Rock", "Slow Rock", "Big Band",
	  "Chorus", "Easy Listening", "Acoustic", "Humour", "Speech", "Chanson", "Opera",
	  "Chamber Music", "Sonata", "Symphony", "Booty Bass", "Primus", "Porn Groove",
	  "Satire", "Slow Jam", "Club", "Tango", "Samba", "Folklore", "Ballad", "Power Ballad",
	  "Rhythmic Soul", "Freestyle", "Duet", "Punk Rock", "Drum Solo", "A Cappella",
	  "Euro-House", "Dance Hall", "Goa", "Drum & Bass", "Club-House", "Hardcore", "Terror",
	  "Indie", "BritPop", "Negerpunk", "Polsk Punk", "Beat", "Christian Gangsta Rap",
	  "Heavy Metal", "Black Metal", "Crossover", "Contemporary Christian",
	  "Christian Rock", "Merengue", "Salsa", "Thrash Metal", "Anime", "JPop", "Synthpop" ]

	def __init__(self, flname = None):
		UserDict.__init__(self)
		self["name"] = flname
		self.__parse()
		self.dir, file = os.path.split(flname)
		self["new_name"] = file
		
	def __parse(self):
		try:
			fl = open(self["name"], "r")
			try:
				fl.seek(-128, 2)
				tagdata = fl.read()
			finally:
				fl.close()
			if tagdata[:3] == 'TAG':
				for tag, (start, end, func) in self.tagDataMap.items():
					self[tag] = func(tagdata[start:end])
			else:
				for tag in self.tagDataMap.keys():
					self[tag] = ''
		except IOError,msg:
			raise BigBooBoo("Error opening file %s\n%s" % (flname, msg))
			
	def __setitem__(self, key, item):
		if key in [ "title", "artist", "album", "comment" ]:
			item = item[:30]			# Truncate upto 30 chars
		elif key == "year":
			item = item[:4]				# Truncate upto 4 chars
		elif key == "genre":
			item = self.getGenre(item)
		self.data[key] = item
	
	def __repr__(self):
		Tag = '\n'.join(['%10s : %s' % (k, v) for k, v in self.items()
		                 if k != 'name' and k != 'new_name'])
		return "Filename: %s\nNew Name(if changed): %s\nID Tag:\n%s" %	\
		       (os.path.basename(self["name"]), self["new_name"], Tag)
		
	def getGenre(self, data):
		if data in self.glist:
			return data
		elif str(data).isdigit():
			try:
				return self.glist[data]
			except IndexError:
				return 'Other'
		return ''
	
	def toggleRulez(self, ruleNos):
		for i in ruleNos:
			self.rulez[i][2][0] = not self.rulez[i][2][0]

	def changeName(self):
		for i in self.rulez:
			if i[2][0]:
				self["new_name"] = re.sub(i[1], i[2][1], self["new_name"])
		self["new_name"] = ' '.join([ x.capitalize() for x in self["new_name"].split()])

	def changeTag(self, tagdata):
		for tag in tagdata.keys():
			self[tag] = tagdata[tag]

	def clearTag(self):
		for tag in self.tagDataMap.keys():
			self[tag] = '' 
			
	def writeTag(self):
		try:
			tag = 'TAG%-30s%-30s%-30s%-4s%-30s%s' % \
				(self["title"], self["artist"], self["album"],
				self["year"], self["comment"], chr(self.glist.index(self["genre"])))
			try:
				fl = open(self["name"], "rb")
				p = fl.read()
				fl.close()
				p = p[:-128]+tag
				fl = open(self["name"], "wb")
				fl.write(p)
			finally:
				fl.close()
		except IOError. msg:
			raise BigBooBoo('Error writing tag for file: %s\n%s' % (self[name], msg))
	
	def commitNameChange(self):
		src = self['name']
		dest = os.path.join(self.dir, self["new_name"])
		os.rename(src, dest)
		self["name"] = dest

def isMp3(file):
	fl = open(file, 'rb')
	if ord(fl.read(1)) & 0xff == 255 and ord(fl.read(1)) | 0x1f == 255:
		return 1
	return 0


def main():
	initial_flist = os.sys.argv[1:] or os.listdir('.')
	names = []
	for file in initial_flist:
		if os.path.isfile(file): # and isMp3(file):# The isMp3() func. is B-A-D
			fl = mp3file(file)                       # either comment it out or
			fl.changeName()                          # write a proper one yourself
			print fl
		else:
			print file, ": not a vaild mp3 file"
		raw_input()

if __name__ == '__main__':
	main()
