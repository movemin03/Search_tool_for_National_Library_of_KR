국립중앙박물관 api를 활용한 도서 정보 검색 프로그램입니다. 대량의 정보를 한번에 검색하기 위해 제작되었습니다.

큰 꿈 작은 도서관 CLIB 도서관리 프로그램에 맞춘 프로그램으로 아래 프로그램과 함께 사용하시면 편리하게 이용하실 수 있습니다.
http://www.smalllibrary.co.kr/xe/download

책의 ISBN 으로 전체 정보를 검색하거나 저자기호만 따로 검색하실 수도 있습니다.

주의사항 및 사용법:
  1. 여러 자료를 검색하실 때 구분자는 "띄어쓰기"입니다. "," 나 "엔터"는 사용하실 수 없습니다. 예를 들어, 9788990364821 8990364825 이런 식으로 입력 바랍니다.
  2. 검색 시 ISBN 13 형식으로 검색하실 것을 권장드립니다만 ISBN 10 형식도 검색은 가능합니다.
  3 엑셀 파일에서 파일을 긁어오실 때 자료가 새로로 입력되어 있다면(혹은 엔터가 쳐져 있다면) 자료가 정상적으로 검색되지 않을 수 있습니다. 엑셀 concatenate 함수와 "& " 를 이용하여 한 셀 내에 자료가 모두 존재하도록 설정해주시고 복사붙여넣기 해주시길 바랍니다.
  4. 저자명은 한글로만 작동합니다. 만약 저자명이 영어로 되어있는 경우, 로마자 표기법에 따라 한글 발음으로 변환 후 입력 바랍니다.
  5. 검색에 성공한 자료는 "성공목록.xlsx"로, 실패한 경우 "실패목록.txt"로 저장됩니다. 실패 목록에 있는 것은 보통 수동으로 일일이 검색해야 합니다.
  6. 최신 파일은 py, txt 파일로 제공합니다. exe 파일은 버전 업데이트가 느립니다. 필요시 pyinstaller 를 통해 exe 파일로 변경 바랍니다.
  exe 파일은 https://github.com/movemin03/Search_tool_for_National_Library_of_KR/blob/master/programfor.exe 여기서 다운로드 받을 수 있습니다.

For Grobal user:

for example:
9788990364821 8990364825

success data will be restore in your desktop with name "성공목록.xlsx"
and fail data will be restore with name "실패목록.txt"
You should search data in your fail list manually.

you can make exe file by Pyinstaller
like "pyinstaller --onefile --hidden-import openxl --noconsole GUI_Search_tool_for_National_Library_of_KR_3456.py"

Or just download "exe" file.

py user:
  tkinter
  urllib
  pandas
  openxl
  requests
  are imported
