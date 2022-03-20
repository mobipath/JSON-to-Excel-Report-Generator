# JSON to Excel Report Generator
python-excel-report-generator is a microservice written in python to write an excel report in xlsx file. It is born from the lack of an existing library to write natively from Python the Office Open XML format. It can be accessed from any other programming framework using post request. 

### Framework and languages
* Django 4.0
* python 3.7
* pandas
* Django Rest Framework
* openpyxl

### How to use this service
**url = "http://excel.iofact.com/api/excel_export"**

**request method = post**

**data type = json**

#### Data preparation
**"topHeader"**: topHeader is a list of dictionary. In topHeader, user must explicitly define the cell information with column index. 

```python
topHeader = [{
    "column": "A11:A13",
    "title": "Cell value",
    "font": {
        "font_size": "11",
        'font_family': "Calibri",
        "bold": True,
        "italic": False,
        "underline": "none",
        "color": "FF000000"
    },
    "alignment": {
        "horizontal": "center",
        "vertical": "center"
    }
}]

```
**"columnHeader"**: columnHeader contains the column information of the excel report. Users can create a dynamic excel report with the option of cell merging using columnHeader. In this key, there is no need to tell the excel column index. Insted children cell (which cell need to be merged) should send in children parameter. However, a key parameter should be sent for mapping the data with the column.  

```python
columnHeader = [
    {'title': 'A', 'key': 'a', 'style': {'font': {'font_size': '11', 'font_family': 'Calibri', 'bold': True, 'italic': False,
                      'underline': 'none', 'color': 'FF000000'},
             'alignment': {'horizontal': 'center', 'vertical': 'center'}}},
    {'title': 'B', 'key': 'b',
     'children':
         [
             {'title': 'C', 'key': 'c'},
             {'title': 'X', 'key': 'x'},
             {'title': 'D', 'key': 'd',
              'children': [
                  {'title': 'E', 'key': 'e'},
                  {'title': 'F', 'key': 'f'}
              ]
              }
         ],
     },
    {
        'title': 'G', 'key': 'g',
        'children': [
             {'title': 'H', 'key': 'h'},
             {'title': 'J', 'key': 'j'},
             {'title': 'I', 'key': 'i',
              'children': [
                  {'title': 'K', 'key': 'k'},
                  {'title': 'L', 'key': 'l',
                   'children': [
                        {'title': 'H', 'key': 'h'},
                        {'title': 'J', 'key': 'j'}]
                   }
              ]
              }
        ]
    },
    {
        'title': 'Z', 'key': 'z'
    }
]
```

**"explicitColumnHeader"**: Explicitly define the cell index. No children paremeter will be accepted. Used method is as same as "topHeader."

```python
explicitColumnHeader = [{
  "column": "A1:A3",
  "title": "Cell information"
}]
```
**"tableData"**: All the cell data will be send in "tableData". 

#### Sample Data preparation (API testing)
* simple Api Data
```json
data = { 
   "explicitColumnHeader":[ 
      { 
         "column":"A1:A3",
         "title":"Budget"
      },
      { 
         "column":"B1:B3",
         "title":"Events"
      }
   ],
   "tableData":{ 
      "Budget":[ 
         10000,
         15000,
         20000
      ],
      "Events":[ 
         "A",
         "B",
         "C"
      ]
   }
}

excelReport = requests.post("http://excel.iofact.com/api/excel_export", json=data)
```


![alt text][sample output]

[sample output]: https://github.com/devnesthq/python-excel-report-generator/blob/master/Example/Sample%20Output
