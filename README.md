# work_automated

Quick Post-Process
by Jake and Jaleel

For all
-=-

1. Place this directory (quickpp) in C:\
2. Make sure directory includes "data_collection_checklist_updates.xlsx" and a folder containing all the scripts under "\scripts\"
3. Open CMD and install (py -m pip install openpyxl)

For batch
-=-

1. Open CMD and install (py -m pip install gspread) and (py -m pip install oauth2client)
2. Copy quickpp_batch.py in the root directory of the raw files you wish to convert
	For example, if you have a folder C:\raw\ that contains 0001, 0002, etc., place quickpp_batch.py in C:\raw
3. Run. Go get yourself a snack.

Do not close until script informs you it has finished. Review output to check for file problems.


For single conversion
-=-

1. Copy quickpp_single.py in the root directory of the raw files you wish to convert (next to asr, wuw, etc.)
2. Run. Enter in proper script and room numbers manually.

Do not close until script informs you it has finished.
