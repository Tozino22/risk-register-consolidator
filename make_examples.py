# -*- coding: utf-8 -*-
"""Generates the example source registers + mapping file in examples/ (run once; committed output)."""
import os
from openpyxl import Workbook

os.makedirs("examples/sources", exist_ok=True)

# --- Example register 1: clean IT register, 5x5 ---
wb = Workbook()
ws = wb.active
ws.append(["ID", "Risk", "Description", "Likelihood", "Impact", "Score", "Owner"])
for row in [
    ["ITS-01", "Ransomware attack", "Ransomware encrypts file servers and halts operations", 4, 5, 20, "IT Manager"],
    ["ITS-02", "Phishing / credential theft", "Staff credentials stolen via phishing emails", 5, 4, 20, "IT Manager"],
    ["ITS-03", "Unpatched systems", "Known vulnerabilities exploited on unpatched servers", 3, 4, 12, "SysAdmin"],
    ["ITS-04", "Lost/stolen laptop", "Laptop with company data lost or stolen", 2, 3, 6, "SysAdmin"],
    ["ITS-05", "Cloud provider outage", "Primary cloud region outage disrupts services", 2, 3, 6, "IT Manager"],
]:
    ws.append(row)
wb.save("examples/sources/IT_Security_Register.xlsx")

# --- Example register 2: messy ops log, High/Med/Low ---
wb = Workbook()
ws = wb.active
ws.append(["#", "Risk", "Rating", "Assigned To", "Status"])
for row in [
    ["OPS-1", "Cyber attack / hackers get into our systems", "High", "Ops Director", "Open"],
    ["OPS-2", "Key supplier fails and we cannot get stock on time", "Medium", "Procurement Lead", "Open"],
    ["OPS-3", "Fire or flood damages the premises", "Low", "Facilities", "Open"],
    ["OPS-4", "Cannot find enough staff during peak season", "High", "HR Manager", "Open"],
    ["OPS-5", "Losing laptops with client data on them", "Medium", "Ops Director", "Open"],
]:
    ws.append(row)
wb.save("examples/sources/Ops_Risk_Log.xlsx")

# --- Example mapping ---
wb = Workbook()
ws = wb.active
ws.title = "mapping"
ws.append(["register_name", "file", "sheet", "header_row", "col_title", "col_description",
           "col_likelihood", "col_impact", "col_rating", "col_owner", "col_ref", "scale"])
ws.append(["IT Security Register", "IT_Security_Register.xlsx", "", 1, "B", "C", "D", "E", "", "G", "A", "5x5"])
ws.append(["Ops Risk Log", "Ops_Risk_Log.xlsx", "", 1, "B", "", "", "", "C", "D", "A", "HML"])
wb.save("examples/mapping_example.xlsx")

print("Examples written to examples/")
