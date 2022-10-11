'Copyright © 2020 Чечкенёв Андрей
'
'This program is free software: you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation, either version 3 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program.  If not, see <https://www.gnu.org/licenses/>.

'Season Wallpaper for Windows
'by Chechkenev Andrew (DarkCat09/CodePicker13)


'USED PICTURES
'###############################################################################################################
'
'Winter:	http://fonday.ru/info/16812-4168121db61.html
'		https://www.rabstol.net/oboi/winter/1288-russkaya-zima.html
'		https://zastavok.net/zima/54902-gory_zima_sneg_eli_vershiny.html
'		https://astrologics.ru/wp-content/uploads/2020/01/Goroskop-na-nedelyu-27-yanvarya-2-fevralya.jpg
'
'Spring:	http://dom-cvety.com/photo/vesna/1472-krasivye-oboi-vesna.html#photo
'		https://nicefon.ru/oboi/vesna_oboi_cveta_radugi_krasota_kartinki_cvety.html
'		http://wallpapers-image.ru/1920x1080/spring/shirokoformatnye-oboi-hd-vesna-1920x1080.php
'		https://mebel-go.ru/photo/9710-skachat_zhivye_oboi_na_rabochiy_stol_vesna.html
'
'Summer:	https://zastavok.net/leto/54677-pole_rozh_nebo_oblaka_goluboe_nebo_leto.html
'		https://nicefon.ru/oboi/derevyya_les_leto_otraghenie_zeleny_priroda_ozero.html
'		https://wallpapers.99px.ru/wallpapers/309335/
'
'Autumn:	https://www.rabstol.net/oboi/gardens/10379-solnechnaya-osen.html
'		http://wallpapers-images.ru/1920x1080/autumn/oboi_osen_1920x1080.php
'		http://wallpapers-image.ru/1920x1080/autumn/oboi-osen-kartinki-autumn-1920x1080.php
'		http://wallpapers-image.ru/1280x1024/autumn/oboi-osen-foto-wallpapers-autumn-1280x1024.php
'		https://www.goodfon.ru/download/voda-klen-listia-osen/2560x1440/
'
'###############################################################################################################

On Error Resume Next
Dim seasonsWp(4)
Dim monthsWp(12)
Dim Wsh, FSO

'USER PREFERENCES
'True - enabled, False - disabled.
'##########################################################################################

'+++++++ MAIN PREFS ++++++

enableWpChanging = True		'Enabling main functionally of this script
backupPreviousWp = True		'Enabling backuping previous wallpaper (not season-wp)
changeWpEveryMonth = False	'Enabling changing wallpaper every month, instead of season

'+++++++++++++++++++++++++


'----- WP FOR SEASONS ----

seasonsWp(0) = -1	'use -1 for random picture.	This is a winter.
seasonsWp(1) =  4	'				This is a spring.
seasonsWp(2) = -1	'				This is a summer.
seasonsWp(3) =  5	'				This is an autumn.

'-------------------------


'***** WP FOR MONTHS *****

'If changeWpEveryMonth is disabled, the following options will not work

monthsWp(11) = "1.jpg"		'wallpaper on December.	I recommend use 1st picture.
monthsWp(0)  = "2.jpg"		'wallpaper on January.	I recommend use 2nd picture for light theme, and 3rd picture for dark theme.
monthsWp(1)  = "4.jpg"		'wallpaper on Febraury.	I recommend use 4th picture.

monthsWp(2)  = "1.jpg"		'wallpaper on March.
monthsWp(3)  = "2.jpg"		'wallpaper on April.	I recommend use 2nd picture.
monthsWp(4)  = "4.jpg"		'wallpaper on May.

monthsWp(5)  = "1.jpg"		'wallpaper on June.	I recommend use 1st picture.
monthsWp(6)  = "2.jpg"		'wallpaper on July.	I recommend use 2nd picture.
monthsWp(7)  = "3.jpg"		'wallpaper on August.	I recommend use 3rd picture.

monthsWp(8)  = "2.jpg"		'wallpaper on September.
monthsWp(9)  = "3.jpg"		'wallpaper on October.
monthsWp(10) = "5.jpg"		'wallpaper on November.	I recommend use 5th picture.

'*************************

'##########################################################################################

Set Wsh = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")
curMonth = Month(Date())

If (curMonth = 12) or (curMonth > 0 and curMonth < 3) then
curSeason = "winter,0"

ElseIf curMonth > 2 and curMonth < 6 then
curSeason = "spring,1"

ElseIf curMonth > 5 and curMonth < 9 then
curSeason = "summer,2"

ElseIf curMonth > 8 and curMonth < 12 then
curSeason = "autumn,3"
End If


If changeWpEveryMonth then

curWp = monthsWp(curMonth-1)

Else

If seasonsWp(CInt(Split(curSeason, ",")(1))) = -1 then
Randomize
curWp = (Int(Rnd*(FSO.GetFolder(Wsh.CurrentDirectory & "\wallpapers\" & Split(curSeason, ",")(0)).Files.Count))+1) & ".jpg"

Else
curWp = seasonsWp(CInt(Split(curSeason, ",")(1))) & ".jpg"
End If
End If

If backupPreviousWp then

prevWp = Wsh.RegRead("HKEY_CURRENT_USER\Control Panel\Desktop\Wallpaper")
FSO.CopyFile prevWp, Wsh.CurrentDirectory & "\wallpapers\backup." & Split(prevWp, ".")(UBound(Split(prevWp, "."))), True
End If

curWp = Wsh.CurrentDirectory & "\wallpapers\" & Split(curSeason, ",")(0) & "\" & curWp
Wsh.RegWrite "HKEY_CURRENT_USER\Control Panel\Desktop\Wallpaper", curWp, "REG_SZ"
Wsh.Run """%SystemRoot%\System32\RUNDLL32.EXE"" user32.dll,UpdatePerUserSystemParameters", 1, True
