from distutils.core import setup
import py2exe

import os, sys, io, time, datetime, serial, serial.tools.list_ports, openpyxl, PIL, string

from datetime import date

from PIL import Image

setup(console=["r12.py"])