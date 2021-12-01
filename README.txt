*BEFORE DEBUGGING AND DEVELOPING*

Need Python 3.9.7 !NOT 3.10, you will not be able to run pyinstaller: 
https://www.python.org/downloads/release/python-3100/

Need to download the C++ Build Tools via Visual Studio installer : 
https://docs.microsoft.com/en-us/cpp/windows/latest-supported-vc-redist?view=msvc-160

pip install numpy
pip install pandas
pip install easygui
pip install openpyxl

pip install pyinstaller

To create a single run application (no need for debug environment...) ** from running directory (in terminal) **
{
    pyinstaller AnalyzeTestOutputwPandas.py
    cd App
    pyinstaller ../AnalyzeTestOutputwPandas.py --noconfirm
    xcopy ..\data .\dist\AnalyzeTestOutputwPandas\data /E/H/I
} 
You can then copy and paste the App/dist/AnalyzeTestOutputwPandas folder to whereever you'd like.
