# Using Python to apply climate change objectives of EU Taxonomy

To read the article, please visit the [developers portal](https://developers.refinitiv.com/en/article-catalog/article/Analyze_EU_Taxonomy_climate_change). The source code for the article is available in the *taxo.py*.

Application to generate the EU Taxonomy for a list of securities. The portfolio input file must contain a column header called RIC in the first sheet.
The taxonomy data is retrieved and calculated for all the instruments in the RIC column.


## Prequisites:
-	A working Eikon installation

-	An Eikon AppKey. This can be generated as follows:
		Logon to the AppKey generator at https://amers1.apps.cp.thomsonreuters.com/apps/AppkeyGenerator, with your Eikon Credentials   
			or   
		Type in APPKEY in the Eikon Search bar   
		Input a valid App Name like "Portfolio Taxonomy"   
		Select "Eikon Data API" checkbox   
		Register the new app   
		Copy the AppKey. This will be required in the python sample   

-	Python 3.6 or higher

-	Refinitiv Eikon module for python
		pip install eikon

-	OpenPyXL module to read/write Excel files
		pip install openpyxl


## Usage:
Usage: 	python taxo.py APP_KEY [-h] [-i INPUT] [-r REPORT]   
Params:   
  APP_KEY = Required, appkey generated using the instructions above   
  INPUT 	= Optional, input portfolio excel file. Default is "input.xlsx"   
  REPORT 	= Optional, output generated excel file. Default is "report.xlsx"   

E.g:   
  python taxo.py __MY_APP_KEY__ -i SP500Port.xlsx -r GeneratedReport.xlsx   
  python taxo.py __MY_APP_KEY__   
