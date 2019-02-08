
del %1report.txt >%1aaa.txt

%3prog\ST-LINK_CLI.exe -c SWD >>%1aaa.txt
%3prog\ST-LINK_CLI.exe -ME >>%1aaa.txt
%3prog\ST-LINK_CLI.exe -P %1%2.hex >>%1aaa.txt
%3prog\ST-LINK_CLI.exe -V >>%1aaa.txt
%3prog\ST-LINK_CLI.exe -Rst >>%1aaa.txt
%3prog\ST-LINK_CLI.exe -Run >>%1aaa.txt

ren %1aaa.txt report.txt



rem il programma si trova qui: Programmi (x86)\STMicroeloectronics\STM32 ST-LINK Utility\ST-LINK Utility\ST-LINK_CLI.exe









