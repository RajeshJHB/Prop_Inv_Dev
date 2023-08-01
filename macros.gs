{\rtf1\ansi\ansicpg1252\cocoartf2639
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fnil\fcharset0 Menlo-Regular;}
{\colortbl;\red255\green255\blue255;\red20\green67\blue174;\red246\green247\blue249;\red46\green49\blue51;
\red186\green6\blue115;\red24\green25\blue27;\red162\green0\blue16;\red18\green115\blue126;}
{\*\expandedcolortbl;;\cssrgb\c9412\c35294\c73725;\cssrgb\c97255\c97647\c98039;\cssrgb\c23529\c25098\c26275;
\cssrgb\c78824\c15294\c52549;\cssrgb\c12549\c12941\c14118;\cssrgb\c70196\c7843\c7059;\cssrgb\c3529\c52157\c56863;}
\margl1440\margr1440\vieww11520\viewh8400\viewkind0
\deftab720
\pard\pardeftab720\partightenfactor0

\f0\fs26 \cf2 \cb3 \expnd0\expndtw0\kerning0
\outl0\strokewidth0 \strokec2 function\cf4 \strokec4  \cf5 \strokec5 SortSTByDate\cf4 \strokec4 () \{\cb1 \
\pard\pardeftab720\partightenfactor0
\cf4 \cb3   \cf2 \strokec2 var\cf4 \strokec4  \cf6 \strokec6 spreadsheet\cf4 \strokec4  = \cf5 \strokec5 SpreadsheetApp\cf4 \strokec4 .\cf6 \strokec6 getActive\cf4 \strokec4 ();\cb1 \
\cb3   \cf6 \strokec6 spreadsheet\cf4 \strokec4 .\cf6 \strokec6 getRange\cf4 \strokec4 (\cf7 \strokec7 'C7:H7'\cf4 \strokec4 ).\cf6 \strokec6 activate\cf4 \strokec4 ();\cb1 \
\cb3   \cf2 \strokec2 var\cf4 \strokec4  \cf6 \strokec6 currentCell\cf4 \strokec4  = \cf6 \strokec6 spreadsheet\cf4 \strokec4 .\cf6 \strokec6 getCurrentCell\cf4 \strokec4 ();\cb1 \
\cb3   \cf6 \strokec6 spreadsheet\cf4 \strokec4 .\cf6 \strokec6 getSelection\cf4 \strokec4 ().\cf6 \strokec6 getNextDataRange\cf4 \strokec4 (\cf5 \strokec5 SpreadsheetApp\cf4 \strokec4 .\cf5 \strokec5 Direction\cf4 \strokec4 .\cf5 \strokec5 DOWN\cf4 \strokec4 ).\cf6 \strokec6 activate\cf4 \strokec4 ();\cb1 \
\cb3   \cf6 \strokec6 currentCell\cf4 \strokec4 .\cf6 \strokec6 activateAsCurrentCell\cf4 \strokec4 ();\cb1 \
\cb3   \cf6 \strokec6 spreadsheet\cf4 \strokec4 .\cf6 \strokec6 getActiveRange\cf4 \strokec4 ().\cf6 \strokec6 offset\cf4 \strokec4 (\cf8 \strokec8 1\cf4 \strokec4 , \cf8 \strokec8 0\cf4 \strokec4 , \cf6 \strokec6 spreadsheet\cf4 \strokec4 .\cf6 \strokec6 getActiveRange\cf4 \strokec4 ().\cf6 \strokec6 getNumRows\cf4 \strokec4 () - \cf8 \strokec8 1\cf4 \strokec4 ).\cf6 \strokec6 sort\cf4 \strokec4 (\{\cf6 \strokec6 column\cf4 \strokec4 : \cf8 \strokec8 3\cf4 \strokec4 , \cf6 \strokec6 ascending\cf4 \strokec4 : \cf2 \strokec2 true\cf4 \strokec4 \});\cb1 \
\cb3   \cf6 \strokec6 spreadsheet\cf4 \strokec4 .\cf6 \strokec6 getRange\cf4 \strokec4 (\cf7 \strokec7 'C7'\cf4 \strokec4 ).\cf6 \strokec6 activate\cf4 \strokec4 ();\cb1 \
\cb3 \};\cb1 \
\
\pard\pardeftab720\partightenfactor0
\cf2 \cb3 \strokec2 function\cf4 \strokec4  \cf6 \strokec6 clearStatement\cf4 \strokec4 () \{\cb1 \
\pard\pardeftab720\partightenfactor0
\cf4 \cb3   \cf2 \strokec2 var\cf4 \strokec4  \cf6 \strokec6 spreadsheet\cf4 \strokec4  = \cf5 \strokec5 SpreadsheetApp\cf4 \strokec4 .\cf6 \strokec6 getActive\cf4 \strokec4 ();\cb1 \
\cb3   \cf6 \strokec6 spreadsheet\cf4 \strokec4 .\cf6 \strokec6 getRange\cf4 \strokec4 (\cf7 \strokec7 'C:H'\cf4 \strokec4 ).\cf6 \strokec6 activate\cf4 \strokec4 ();\cb1 \
\cb3   \cf6 \strokec6 spreadsheet\cf4 \strokec4 .\cf6 \strokec6 getActiveRangeList\cf4 \strokec4 ().\cf6 \strokec6 clear\cf4 \strokec4 (\{\cf6 \strokec6 contentsOnly\cf4 \strokec4 : \cf2 \strokec2 true\cf4 \strokec4 , \cf6 \strokec6 skipFilteredRows\cf4 \strokec4 : \cf2 \strokec2 true\cf4 \strokec4 \});\cb1 \
\cb3 \};\cb1 \
}