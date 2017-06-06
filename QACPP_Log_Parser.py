# -*- coding: utf-8 -*-
"""
Copyright 2015 Joohyun Lee(ppiazi@gmail.com)

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
"""
import sys
import os
import getopt
import xlsxwriter

__author__ = 'ppiazi'
__version__ = 'v0.0.1'

LOG_PTN_QAC_OUT = "(prqa) OUT: qac-gui"
LOG_PTN_QAC_IN  = "(prqa) IN: qac-gui"
LOG_PTN_QACPP_OUT = "(prqa) OUT: qacpp-gui"
LOG_PTN_QACPP_IN  = "(prqa) IN: qacpp-gui"
LOG_PTN_DENIED = "(prqa) DENIED:"

EXCEL_COLS = ["No", "TYPE", "OUT_DATE", "OUT_TIME", "IN_TIME", "BY", "ORIGINAL_LOG"]

def print_log_found(i, t):
    print("%04d : %s (%s to %s) by %s" % (i, t["type"], t["out_time"], t["in_time"], t["by"]))

def analyze_qac_log(log_file):

    try:
        f = open(log_file, "r")
    except Exception as e:
        print("Exception : %s" % (str(e)))

    lines = f.readlines()

    qac_log_list = []
    i = 0
    j = 0
    for each_line in lines:
        j = j + 1
        if LOG_PTN_QAC_OUT in each_line:
            i = i + 1
            t = analyze_qac_gui(each_line, lines, j)
            qac_log_list.append(t)
            print_log_found(i, t)
            continue
        elif LOG_PTN_QACPP_OUT in each_line:
            i = i + 1
            t = analyze_qacpp_gui(each_line, lines, j)
            qac_log_list.append(t)
            print_log_found(i, t)
            continue
        elif LOG_PTN_DENIED in each_line:
            i = i + 1
            t = analyze_denied(each_line)
            qac_log_list.append(t)
            print_log_found(i, t)
            continue

    save_as_excel(log_file, qac_log_list)

def save_as_excel(log_file, qac_log_list):
    excel_file_name = log_file + ".xlsx"

    wbk = xlsxwriter.Workbook(excel_file_name)
    sheet = wbk.add_worksheet("QAC_LOG")

    i = 0
    for col in EXCEL_COLS:
        sheet.write(0, i, col)
        i = i + 1

    i = 1
    for row in qac_log_list:
        sheet.write(i, 0, i)
        sheet.write(i, 1, row["type"])
        sheet.write(i, 2, row["out_date"])
        sheet.write(i, 3, row["out_time"])
        sheet.write(i, 4, row["in_time"])
        sheet.write(i, 5, row["by"])
        sheet.write(i, 6, row["origin_log"])
        i = i + 1

    wbk.close()

def analyze_denied(line):
    denied_log_time, denied_to = parse_denied_log(line)

    t = {}
    t["by"] = denied_to
    if "qacpp" in line:
        t["type"] = "qacpp_denied"
    elif "qac" in line:
        t["type"] = "qac_denied"
    else:
        t["type"] = "etc_denied"

    t["out_time"] = denied_log_time
    t["in_time"] = denied_log_time
    t["out_date"] = denied_log_time[:5]
    t["origin_log"] = line

    return t

def analyze_qac_gui(each_line, lines, j):
    out_log_time, out_log_by = parse_out_log(each_line)

    in_log_time = None
    # 그 이후의 IN 로그를 찾는다.
    for find_in_log in lines[j:]:
        if LOG_PTN_QAC_IN in find_in_log:
            t_in_log_time, t_in_log_by = parse_out_log(find_in_log)
            if t_in_log_by == out_log_by:
                in_log_time = t_in_log_time
                break
    t = {}
    t["by"] = out_log_by
    t["type"] = "qac_out_in"
    t["out_time"] = out_log_time
    t["out_date"] = out_log_time[:5]

    if in_log_time != None:
        t["in_time"] = in_log_time
    else:
        t["in_time"] = "unknown"
    t["origin_log"] = each_line

    return t

def analyze_qacpp_gui(each_line, lines, j):
    out_log_time, out_log_by = parse_out_log(each_line)

    in_log_time = None
    # 그 이후의 IN 로그를 찾는다.
    for find_in_log in lines[j:]:
        if LOG_PTN_QACPP_IN in find_in_log:
            t_in_log_time, t_in_log_by = parse_out_log(find_in_log)
            if t_in_log_by == out_log_by:
                in_log_time = t_in_log_time
                break
    t = {}
    t["by"] = out_log_by
    t["type"] = "qacpp_out_in"
    t["out_time"] = out_log_time
    t["out_date"] = out_log_time[:5]

    if in_log_time != None:
        t["in_time"] = in_log_time
    else:
        t["in_time"] = "unknown"
    t["origin_log"] = each_line

    return t

def parse_denied_log(denied_log):
    to = ""
    denied_time = ""

    denied_time = denied_log[:11]
    to_inx = denied_log.find("to")
    to = denied_log[to_inx:].split(" ")[1]

    return denied_time, to

def parse_out_log(out_log):
    by = ""
    out_in_time = ""

    out_in_time = out_log[:11]
    by_inx = out_log.find("by")
    by = out_log[by_inx:].split(" ")[1]

    return out_in_time, by

def printUsage():
    print("QAC_Log_Parser.py [-f <file>]")
    print("    Version %s" % __version__)
    print("    Options:")
    print("    -f : set a target log file")

if __name__ == "__main__":
    optlist, args = getopt.getopt(sys.argv[1:], "f:")

    p_target_file = None

    for op, p in optlist:
        if op == "-f":
            p_target_file = p
        else:
            print("Invalid Argument : %s / %s" % (op, p))

    if p_target_file == None:
        printUsage()
        os._exit(1)

    analyze_qac_log(p_target_file)

    analyze_qac_log("QAC_QACPP.log")