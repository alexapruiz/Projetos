#!/usr/bin/python

##########################################################
#
#  FREESCO Dialer Client
#
#  Updated by Martin Stevens - budgester@budgester.com
#       
#  Compatability for Freesco 0.27 added - 28/12/2000
#
##########################################################

# web admin server must be running on port 82 for this to
# work.  

SERVER="*.*.*.*" # address of gateway machine
USER="*****"  # probably admin for you or make a new account
PASS="*****"  # users password

# don't change below this line ###########################

from Tkinter import *
from urllib import *
import string
import sys
import time

def setStatus(sts):
    """
    set dialup status. 0=hangup, 1=dial
    """

    if (sts==0):
        fl=urlopen("http://" + USER + ":" + PASS + "@" + SERVER + ":82/cgi/usr.cgi?block")
    else:
        OP=""
        fl=urlopen("http://" + USER + ":" + PASS + "@" + SERVER + ":82/cgi/usr.cgi?unblock")
        fl=urlopen("http://" + USER + ":" + PASS + "@" + SERVER + ":82/cgi/usr.cgi?up")

root=Tk()
root.title("FREESCO Dialer")
Button(root, text="Dial", width=10, command=lambda: setStatus(1)).grid(column=1,row=0)
Button(root, text="Hangup", width=10, command=lambda: setStatus(0)).grid(column=2,row=0)
Button(root, text="Quit", width=10, command=root.destroy).grid(column=3,row=0)
mainloop()