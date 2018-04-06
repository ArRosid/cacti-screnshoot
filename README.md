<h1>Cacti Graph Screenshoot</h1>
This is a script to take screenshoot from Cacti graph and export the result to excel file

<h3>Setup guide</h3>
To run this script, you need to install some python libary
<br>
First python library that you need to install is Pillow. For more information about Pillow library, visit the documentation site https://pillow.readthedocs.io/en/5.1.x/. To install this library, use this command
<pre>pip3 install Pillow</pre>
<br>

Second library is xlsxwriter. This library will export the result of screenshoot to excel file. For more information about this library, visit the documentation site http://xlsxwriter.readthedocs.io/. To install this library, use this command
<pre>pip3 install xlsxwriter</pre>
<br>

Third library is selenium, for more information about this library, visit the documentation site http://selenium-python.readthedocs.io/index.html. To install this library, use this command
<pre>pip3 install selenium</pre>
<br>

Last, you need to install chromedriver. To instal chromedriver, you can follow tutorial on this site https://gist.github.com/ziadoz/3e8ab7e944d02fe872c3454d17af31a5
<br>
<h3>Run the script</h3>
When you have installed all of the prerequisite library, you can run the script by this command
<pre>python3 CactiScreenshoot.py</pre>
