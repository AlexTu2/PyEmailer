@echo off
cls
python "C:Github\Python Email\emailer.py" --mail_list %1
timeout 55