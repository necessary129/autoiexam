# autoIExam: Automate trivial IExam tasks that shouldn't have been needed.

## Prerequisites

You should have Python 3 installed. (3.8 recommended. Only tested in 3.8 but 3.6+ should work.)

You should have maayalam added as a language in your OS settings.

You should have `selenium` and `xlrd` installed.

Install them:

```
pip install selenium xlrd --user
```

You should have the latest version of firefox installed.


## Howto

First, Create a custom report in Sampoorna for your class and division only, with `Admission number` as the first column, `Full name (malayalam)` as the second column, `House name` as the third column, `Street/Place` as the fourth column, `Post office` as the fifth column, `PIN Code` as the sixth column and `Revenue district` as the final, seventh column.

Click export csv and save that file to this directory as `report.xls`.

Edit `config.json` and put your username and password in there and then run `autoiexam.py` by double clicking or typing 

```
python3 autoiexam.py
```

In a terminal opened in the directory.
