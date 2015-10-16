# -*- coding: utf-8 -*-
import codecs
import win32com.client
from input import STA_FILE_NAME, GRD_FILE_NAME

SECTIONS = {}

acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument


def get_num(s):
    return float("".join(char for char in s.replace(",", ".") if char.isdigit() or char in ["-", "+", "."]))


def getSecSta(doc):
    f = None
    try:
        f = open(STA_FILE_NAME, "r")
        for line in f:
            section, station = line.strip().split(",")
            if section and station:
                section = section.strip()
                station = float(station)
                SECTIONS[section] = station
        f.close()

        section = doc.Utility.GetString(False, "Input section name: ")
        if not section:
            return None, None

        section = codecs.encode(section, "cp1253")
        section = section.strip()
        try:
            return section, SECTIONS[section]
        except KeyError:
            station = doc.Utility.GetString(False, "Section not found in file, input station value: ")
            if not station:
                return None, None
            return section, float(station)

    except IOError:
        print "Station file not found!"
        s = doc.Utility.GetString(False, "Input section name and station (comma seperated): ")
        if not s:
            return None, None

        section, station = s.strip().split(",")
        section = codecs.encode(section, "cp1253")

        section = section.strip()
        station = float(station)
        return section, station

f = open(GRD_FILE_NAME, "a")
while True:
    section, station = getSecSta(doc)
    if section is None and station is None:
        break

    print "Getting data for section %s at station %.2f" % (section, station)
    polyline, point_clicked = doc.Utility.GetEntity(None, None, Prompt="Select a polyline:")

    origin = doc.Utility.GetPoint(Prompt="Select section origin:")
    originHeight, point_clicked = doc.Utility.GetEntity(None, None, Prompt="Select text that includes origin height:")
    originHeight = get_num(originHeight.TextString)

    pointlist = []
    for i in range(len(polyline.Coordinates)):
        if i % 2 == 0:
            x = polyline.Coordinates[i]
        elif i % 2 == 1:
            y = polyline.Coordinates[i]
            pointlist.append((x, y))

    f.write("*\n")
    f.write("%s      %.2f\n" % (section, station))
    for pnt in pointlist:
        offset = pnt[0] - origin[0]
        height = pnt[1] - origin[1] + originHeight
        f.write("%.2f  %.2f\n" % (offset, height))

f.write("*\n")
f.close()
