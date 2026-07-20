echo Dim obj                                                                                                 >%TEMP%\test-com.vbs
echo Set obj = WScript.CreateObject("DSOFile.OleDocumentProperties")                                        >>%TEMP%\test-com.vbs
echo If Not IsObject(obj) Then                                                                              >>%TEMP%\test-com.vbs
echo 	Wscript.Echo("COM use of DSOFile.OleDocumentProperties failed.")                                    >>%TEMP%\test-com.vbs
echo 	WScript.Quit 1				                                                                        >>%TEMP%\test-com.vbs
echo Else                                                                                                   >>%TEMP%\test-com.vbs
echo 	Wscript.Echo("COM use of DSOFile.OleDocumentProperties ok.")				                        >>%TEMP%\test-com.vbs
echo End If                                                                                                 >>%TEMP%\test-com.vbs
echo Set obj = Nothing                                                                                      >>%TEMP%\test-com.vbs
echo                                                                                                        >>%TEMP%\test-com.vbs
echo Set obj = WScript.CreateObject("MT_PropertyExchange.Globals")                                          >>%TEMP%\test-com.vbs
echo If Not IsObject(obj) Then                                                                              >>%TEMP%\test-com.vbs
echo 	Wscript.Echo("COM use of MT_PropertyExchange.Globals failed.")                                      >>%TEMP%\test-com.vbs
echo 	WScript.Quit 1                                                                                      >>%TEMP%\test-com.vbs
echo Else                                                                                                   >>%TEMP%\test-com.vbs
echo 	Wscript.Echo("COM use of MT_PropertyExchange.Globals ok." & vbcrlf & "DLL Version: " & obj.Version) >>%TEMP%\test-com.vbs
echo End If                                                                                                 >>%TEMP%\test-com.vbs
echo Set obj = Nothing                                                                                      >>%TEMP%\test-com.vbs
echo WScript.Quit 0                                                                                         >>%TEMP%\test-com.vbs
