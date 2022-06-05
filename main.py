from time import sleep

import xlrd
import xlwt

# open the excel file
workbook = xlrd.open_workbook('0.Data_Khen_CN_2122.xls')
# get the first sheet
sheet = workbook.sheet_by_index(0)
# get cell value to data list
data = []
for i in range(sheet.nrows):
    data.append(sheet.row_values(i))

data.pop(0)

for i in range(len(data)):
    data[i].pop(3)
    data[i].pop(3)
    data[i].pop(3)
    data[i][1], data[i][2] = data[i][2], data[i][1]
    data[i][0] = i + 1
#  [TT] [họ tên] [lớp] [danh hiệu]
with open('index.html', 'w', encoding="utf-8") as f:
    f.write('''<!DOCTYPE html>
        <html lang="en">

        <head>
        <meta charset="UTF-8">
        <meta http-equiv="X-UA-Compatible" content="IE=edge">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Credits Bảng Vàng Trường THPT Nguyễn Bỉnh Khiêm</title>
        <link rel="stylesheet" href="css/main.css">
        </head>

        <body style="background-image: url(./images/280247631_490399796174125_6882678655106843732_n.png);">
            <div class='wrapper'>
            <div class='movie'>''')
    for i in range(len(data)):
        f.write(f"<div class='name'>  {data[i][1]} </div>\n")
        f.write(f"<div class='status'>")
        f.write(f"<span class='class'> {data[i][2]} </span> -\n")
        f.write(f"<span class='rank'> Danh hiệu học sinh {data[i][3]} </span>\n")
        f.write(f"</div>")
    f.write("</body></html>")
with open('css/main.css', 'w', encoding="utf-8") as f:
    f.write('''
            /* GRADIEND */
        .ap {
          position: fixed;
          right: 0;
          bottom: 0;
          left: 0;
          height: 40px;
          margin: auto;
          user-select: none;
          color: #333;
          background: rgb(194, 48, 48);
          border-top: 50px solid #f0f;
          z-index: 9999;
        }
        
        /*END*/
        
        @import url('https://fonts.googleapis.com/css2?family=Open+Sans:wght@800&display=swap');
        
        * {
          margin: 0;
          padding: 0;
          box-sizing: border-box;
        }
        
        html,
        body {
          height: 100%;
        }
        
        body {
          background-color: #333;
          background-size: 100% 100%;
          background-position: center;
          overflow: hidden;
          color: white;
          text-transform: uppercase;
        }
        
        .wrapper {
          position: absolute;
          top: 100%;
          left: 50%;
          width: 1000px;
          margin-left: -500px;
          font-family: 'Open Sans', sans-serif;
          text-align: center;
          text-transform: uppercase;
          color: rgb(255, 0, 0);
          animation: 1000s credits linear;
        }
        
        .movie {
          margin-bottom: 50px;
          font-size: 50px;
        }
        
        .status {
          text-align: center;
          font-size: 42px;
          margin-bottom: 8rem;
        }
        
        .name {
          text-align: center;
          margin-bottom: 0rem;
          font-size: 64px;
        }
        
        @keyframes credits {
          0% {
            top: 100%;
          }
        
          100% {
          
            top:''')
    f.write(f"-{round(len(data) * 35.180214723926380368098159509202)}%")
    f.write('''}
    }''')
print('Done')
