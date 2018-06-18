{\rtf1\ansi\ansicpg1252\cocoartf1561\cocoasubrtf100
{\fonttbl\f0\fnil\fcharset134 STSongti-SC-Light;\f1\fnil\fcharset134 STSongti-SC-Regular;}
{\colortbl;\red255\green255\blue255;\red109\green109\blue109;\red255\green255\blue255;\red0\green0\blue109;
\red0\green0\blue0;\red160\green0\blue163;\red128\green63\blue122;\red15\green112\blue1;\red0\green0\blue233;
\red0\green0\blue255;}
{\*\expandedcolortbl;;\cssrgb\c50196\c50196\c50196;\cssrgb\c100000\c100000\c100000;\cssrgb\c0\c0\c50196;
\cssrgb\c0\c0\c0;\cssrgb\c69804\c0\c69804;\cssrgb\c58039\c33333\c55294;\cssrgb\c0\c50196\c0;\cssrgb\c0\c0\c93333;
\cssrgb\c0\c0\c100000;}
\margl1440\margr1440\vieww10800\viewh8400\viewkind0
\deftab720
\pard\pardeftab720\sl440\partightenfactor0

\f0\fs32 \cf2 \cb3 \expnd0\expndtw0\kerning0
\outl0\strokewidth0 \strokec2 # coding=utf-8\cb1 \
\pard\pardeftab720\sl440\partightenfactor0

\f1\b \cf4 \cb3 \strokec4 import 
\f0\b0 \cf5 \strokec5 platform\cb1 \

\f1\b \cf4 \cb3 \strokec4 import 
\f0\b0 \cf5 \strokec5 os, re\cb1 \

\f1\b \cf4 \cb3 \strokec4 import 
\f0\b0 \cf5 \strokec5 sys\cb1 \

\f1\b \cf4 \cb3 \strokec4 import 
\f0\b0 \cf5 \strokec5 xlsxwriter\cb1 \
\

\f1\b \cf4 \cb3 \strokec4 class 
\f0\b0 \cf5 \strokec5 MediaInfo(\cf4 \strokec4 object\cf5 \strokec5 ):\cb1 \
\pard\pardeftab720\sl440\partightenfactor0
\cf5 \cb3     
\f1\b \cf4 \strokec4 def 
\f0\b0 \cf6 \strokec6 __init__\cf5 \strokec5 (\cf7 \strokec7 self\cf5 \strokec5 ):\cb1 \
\cb3         \cf7 \strokec7 self\cf5 \strokec5 .general = []\cb1 \
\cb3         \cf7 \strokec7 self\cf5 \strokec5 .video = []\cb1 \
\cb3         \cf7 \strokec7 self\cf5 \strokec5 .audio = []\cb1 \
\cb3     
\f1\b \cf4 \strokec4 def 
\f0\b0 \cf5 \strokec5 getListByName(\cf7 \strokec7 self\cf5 \strokec5 , name):\cb1 \
\cb3         
\f1\b \cf4 \strokec4 if 
\f0\b0 \cf5 \strokec5 name == 
\f1\b \cf8 \strokec8 'General'
\f0\b0 \cf5 \strokec5 :\cb1 \
\cb3             
\f1\b \cf4 \strokec4 return 
\f0\b0 \cf7 \strokec7 self\cf5 \strokec5 .general\cb1 \
\cb3         
\f1\b \cf4 \strokec4 elif 
\f0\b0 \cf5 \strokec5 name == 
\f1\b \cf8 \strokec8 'Video'
\f0\b0 \cf5 \strokec5 :\cb1 \
\cb3             
\f1\b \cf4 \strokec4 return 
\f0\b0 \cf7 \strokec7 self\cf5 \strokec5 .video\cb1 \
\cb3         
\f1\b \cf4 \strokec4 elif 
\f0\b0 \cf5 \strokec5 name == 
\f1\b \cf8 \strokec8 'Audio'
\f0\b0 \cf5 \strokec5 :\cb1 \
\cb3             
\f1\b \cf4 \strokec4 return 
\f0\b0 \cf7 \strokec7 self\cf5 \strokec5 .audio\cb1 \
\cb3         
\f1\b \cf4 \strokec4 else
\f0\b0 \cf5 \strokec5 :\cb1 \
\cb3             
\f1\b \cf4 \strokec4 return 
\f0\b0 None\cb1 \
\
\cf5 \cb3 \strokec5 mediaInfos = []\cb1 \
\
\cb3 cmd = 
\f1\b \cf8 \strokec8 '{\field{\*\fldinst{HYPERLINK "c:/Users/viclv/Downloads/Compressed/MediaInfo_CLI_0.7.98_Windows_x64/MediaInfo.exe"}}{\fldrslt \cf9 \ul \ulc9 \strokec9 C:/Users/viclv/Downloads/Compressed/MediaInfo_CLI_0.7.98_Windows_x64/MediaInfo.exe}} "%s"' 
\f0\b0 \cf5 \strokec5 % 
\f1\b \cf8 \strokec8 '{\field{\*\fldinst{HYPERLINK "d:/BaiduYunDownload/1.mp4"}}{\fldrslt \cf9 \ul \ulc9 \strokec9 D:/BaiduYunDownload/1.mp4}}'\cb1 \

\f0\b0 \cf5 \cb3 \strokec5 p = os.popen(cmd)\cb1 \
\cb3 lines = p.read().split(
\f1\b \cf8 \strokec8 '\cf4 \strokec4 \\n\cf8 \strokec8 '
\f0\b0 \cf5 \strokec5 )\cb1 \
\cb3 mediaInfo = MediaInfo()\cb1 \
\cb3 types = [
\f1\b \cf8 \strokec8 'General'
\f0\b0 \cf5 \strokec5 , 
\f1\b \cf8 \strokec8 'Video'
\f0\b0 \cf5 \strokec5 , 
\f1\b \cf8 \strokec8 'Audio'
\f0\b0 \cf5 \strokec5 ]\cb1 \
\
\pard\pardeftab720\sl440\partightenfactor0

\f1\b \cf4 \cb3 \strokec4 for 
\f0\b0 \cf5 \strokec5 line 
\f1\b \cf4 \strokec4 in 
\f0\b0 \cf5 \strokec5 lines:\cb1 \
\pard\pardeftab720\sl440\partightenfactor0
\cf5 \cb3     
\f1\b \cf4 \strokec4 if 
\f0\b0 len\cf5 \strokec5 (line.strip()) == \cf10 \strokec10 0\cf5 \strokec5 :\cb1 \
\cb3         
\f1\b \cf4 \strokec4 continue\cb1 \
\pard\pardeftab720\sl440\partightenfactor0
\cf4 \cb3     if 
\f0\b0 \cf5 \strokec5 line.strip() 
\f1\b \cf4 \strokec4 in 
\f0\b0 \cf5 \strokec5 types:\cb1 \
\pard\pardeftab720\sl440\partightenfactor0
\cf5 \cb3         tmpList = mediaInfo.getListByName(line.strip())\cb1 \
\cb3     
\f1\b \cf4 \strokec4 else
\f0\b0 \cf5 \strokec5 :\cb1 \
\cb3         
\f1\b \cf4 \strokec4 if 
\f0\b0 \cf5 \strokec5 tmpList == \cf4 \strokec4 None\cf5 \strokec5 :\cb1 \
\cb3             
\f1\b \cf4 \strokec4 continue\cb1 \
\pard\pardeftab720\sl440\partightenfactor0
\cf4 \cb3         
\f0\b0 \cf5 \strokec5 columnIndex = line.index(
\f1\b \cf8 \strokec8 ':'
\f0\b0 \cf5 \strokec5 )\cb1 \
\pard\pardeftab720\sl440\partightenfactor0
\cf5 \cb3         key, value = line[:columnIndex].strip(), line[columnIndex + \cf10 \strokec10 1\cf5 \strokec5 :].strip()\cb1 \
\cb3         tmpList.append((key, value))\cb1 \
\
\cb3 mediaInfos.append(mediaInfo)\cb1 \
\
\pard\pardeftab720\sl440\partightenfactor0

\f1\b \cf4 \cb3 \strokec4 print
\f0\b0 \cf5 \strokec5 (mediaInfo.getListByName(
\f1\b \cf8 \strokec8 'General'
\f0\b0 \cf5 \strokec5 ))\cb1 \

\f1\b \cf4 \cb3 \strokec4 print
\f0\b0 \cf5 \strokec5 (mediaInfo.getListByName(
\f1\b \cf8 \strokec8 'Video'
\f0\b0 \cf5 \strokec5 ))\cb1 \

\f1\b \cf4 \cb3 \strokec4 print
\f0\b0 \cf5 \strokec5 (mediaInfo.getListByName(
\f1\b \cf8 \strokec8 'Audio'
\f0\b0 \cf5 \strokec5 ))\cb1 \
\
\pard\pardeftab720\sl440\partightenfactor0
\cf5 \cb3 workbook = xlsxwriter.Workbook(
\f1\b \cf8 \strokec8 'VideoInfo.xlsx'
\f0\b0 \cf5 \strokec5 )\cb1 \
\cb3 sheet = workbook.add_worksheet(
\f1\b \cf8 \strokec8 'info'
\f0\b0 \cf5 \strokec5 )\cb1 \
\
\pard\pardeftab720\sl440\partightenfactor0
\cf2 \cb3 \strokec2 #first row\cb1 \
\pard\pardeftab720\sl440\partightenfactor0
\cf5 \cb3 \strokec5 row = \cf10 \strokec10 0\cb1 \
\cf5 \cb3 \strokec5 mediaInfo = mediaInfos[\cf10 \strokec10 0\cf5 \strokec5 ]\cb1 \
\cb3 col = \cf10 \strokec10 0\cb1 \
\cf5 \cb3 \strokec5 sheet.write(row, col, 
\f1\b \cf8 \strokec8 'NO'
\f0\b0 \cf5 \strokec5 )\cb1 \
\pard\pardeftab720\sl440\partightenfactor0

\f1\b \cf4 \cb3 \strokec4 for 
\f0\b0 \cf5 \strokec5 info 
\f1\b \cf4 \strokec4 in 
\f0\b0 \cf5 \strokec5 mediaInfo.general:\cb1 \
\pard\pardeftab720\sl440\partightenfactor0
\cf5 \cb3     col += \cf10 \strokec10 1\cb1 \
\pard\pardeftab720\sl440\partightenfactor0
\cf10 \cb3     \cf5 \strokec5 sheet.write(row, col, info[\cf10 \strokec10 0\cf5 \strokec5 ],workbook.add_format(\{
\f1\b \cf8 \strokec8 'bg_color'
\f0\b0 \cf5 \strokec5 :
\f1\b \cf8 \strokec8 'red'
\f0\b0 \cf5 \strokec5 \}))\cb1 \
\pard\pardeftab720\sl440\partightenfactor0

\f1\b \cf4 \cb3 \strokec4 for 
\f0\b0 \cf5 \strokec5 info 
\f1\b \cf4 \strokec4 in 
\f0\b0 \cf5 \strokec5 mediaInfo.video:\cb1 \
\pard\pardeftab720\sl440\partightenfactor0
\cf5 \cb3     col += \cf10 \strokec10 1\cb1 \
\pard\pardeftab720\sl440\partightenfactor0
\cf10 \cb3     \cf5 \strokec5 sheet.write(row, col, info[\cf10 \strokec10 0\cf5 \strokec5 ],workbook.add_format(\{
\f1\b \cf8 \strokec8 'bg_color'
\f0\b0 \cf5 \strokec5 :
\f1\b \cf8 \strokec8 'yellow'
\f0\b0 \cf5 \strokec5 \}))\cb1 \
\pard\pardeftab720\sl440\partightenfactor0

\f1\b \cf4 \cb3 \strokec4 for 
\f0\b0 \cf5 \strokec5 info 
\f1\b \cf4 \strokec4 in 
\f0\b0 \cf5 \strokec5 mediaInfo.audio:\cb1 \
\pard\pardeftab720\sl440\partightenfactor0
\cf5 \cb3     col += \cf10 \strokec10 1\cb1 \
\pard\pardeftab720\sl440\partightenfactor0
\cf10 \cb3     \cf5 \strokec5 sheet.write(row, col, info[\cf10 \strokec10 0\cf5 \strokec5 ],workbook.add_format(\{
\f1\b \cf8 \strokec8 'bg_color'
\f0\b0 \cf5 \strokec5 :
\f1\b \cf8 \strokec8 'navy'
\f0\b0 \cf5 \strokec5 \}))\cb1 \
\
\pard\pardeftab720\sl440\partightenfactor0
\cf2 \cb3 \strokec2 # info data\cb1 \
\pard\pardeftab720\sl440\partightenfactor0
\cf5 \cb3 \strokec5 row += \cf10 \strokec10 1\cb1 \
\pard\pardeftab720\sl440\partightenfactor0

\f1\b \cf4 \cb3 \strokec4 for 
\f0\b0 \cf5 \strokec5 mediaInfo 
\f1\b \cf4 \strokec4 in 
\f0\b0 \cf5 \strokec5 mediaInfos:\cb1 \
\pard\pardeftab720\sl440\partightenfactor0
\cf5 \cb3     col = \cf10 \strokec10 0\cb1 \
\pard\pardeftab720\sl440\partightenfactor0
\cf10 \cb3     \cf5 \strokec5 sheet.write(row, col, \cf4 \strokec4 str\cf5 \strokec5 (row))\cb1 \
\pard\pardeftab720\sl440\partightenfactor0
\cf5 \cb3     
\f1\b \cf4 \strokec4 for 
\f0\b0 \cf5 \strokec5 info 
\f1\b \cf4 \strokec4 in 
\f0\b0 \cf5 \strokec5 mediaInfo.general:\cb1 \
\cb3         col += \cf10 \strokec10 1\cb1 \
\pard\pardeftab720\sl440\partightenfactor0
\cf10 \cb3         \cf5 \strokec5 sheet.write(row, col, info[\cf10 \strokec10 1\cf5 \strokec5 ])\cb1 \
\pard\pardeftab720\sl440\partightenfactor0
\cf5 \cb3     
\f1\b \cf4 \strokec4 for 
\f0\b0 \cf5 \strokec5 info 
\f1\b \cf4 \strokec4 in 
\f0\b0 \cf5 \strokec5 mediaInfo.video:\cb1 \
\cb3         col += \cf10 \strokec10 1\cb1 \
\pard\pardeftab720\sl440\partightenfactor0
\cf10 \cb3         \cf5 \strokec5 sheet.write(row, col, info[\cf10 \strokec10 1\cf5 \strokec5 ])\cb1 \
\pard\pardeftab720\sl440\partightenfactor0
\cf5 \cb3     
\f1\b \cf4 \strokec4 for 
\f0\b0 \cf5 \strokec5 info 
\f1\b \cf4 \strokec4 in 
\f0\b0 \cf5 \strokec5 mediaInfo.audio:\cb1 \
\cb3         col += \cf10 \strokec10 1\cb1 \
\pard\pardeftab720\sl440\partightenfactor0
\cf10 \cb3         \cf5 \strokec5 sheet.write(row, col, info[\cf10 \strokec10 1\cf5 \strokec5 ])\cb1 \
\
\pard\pardeftab720\sl440\partightenfactor0
\cf5 \cb3 workbook.close()\cb1 \
\
}