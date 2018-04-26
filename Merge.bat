@echo off
set classpath= %classpath%;.\poi-bin-3.14\poi-3.14\poi-3.14-20160307.jar
set classpath= %classpath%;.\poi-bin-3.14\poi-3.14\poi-examples-3.14-20160307.jar
set classpath= %classpath%;.\poi-bin-3.14\poi-3.14\poi-excelant-3.14-20160307.jar
set classpath= %classpath%;.\poi-bin-3.14\poi-3.14\poi-ooxml-3.14-20160307.jar
set classpath= %classpath%;.\poi-bin-3.14\poi-3.14\poi-ooxml-schemas-3.14-20160307.jar
set classpath= %classpath%;.\poi-bin-3.14\poi-3.14\poi-scratchpad-3.14-20160307.jar
set classpath= %classpath%;.\poi-bin-3.14\poi-3.14\ooxml-lib\curvesapi-1.03.jar
set classpath= %classpath%;.\poi-bin-3.14\poi-3.14\ooxml-lib\xmlbeans-2.6.0.jar
set classpath= %classpath%;.\poi-bin-3.14\poi-3.14\lib\commons-codec-1.10.jar
set classpath= %classpath%;.\poi-bin-3.14\poi-3.14\lib\commons-logging-1.2.jar
set classpath= %classpath%;.\poi-bin-3.14\poi-3.14\lib\junit-4.12.jar
set classpath= %classpath%;.\poi-bin-3.14\poi-3.14\lib\log4j-1.2.17.jar
set classpath= %classpath%;.
java MergeWhat F:\sourceData
pause
