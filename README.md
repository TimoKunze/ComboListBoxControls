<p align=center>
  <a href="https://github.com/TimoKunze/ComboListBoxControls/releases/tag/1.5.5">
    <img alt="Release 1.5.5 Release" src="https://img.shields.io/badge/release-1.5.5-0688CB.svg">
  </a>
  <a href="https://github.com/TimoKunze/ComboListBoxControls/releases">
    <img alt="Download ComboListBoxControls" src="https://img.shields.io/badge/download-latest-0688CB.svg">
  </a>
  <a href="https://github.com/TimoKunze/ComboListBoxControls/blob/master/LICENSE">
    <img alt="License" src="https://img.shields.io/badge/license-MIT-0688CB.svg">
  </a>
  <a href="https://www.paypal.com/xclick/business=TKunze71216%40gmx.de&item_name=ComboListBoxControls&no_shipping=1&tax=0&currency_code=EUR">
    <img alt="Donate" src="https://img.shields.io/badge/%24-donate-E44E4A.svg">
  </a>
</p>

# ComboListBoxControls
An ActiveX control library for Visual Basic 6 that contains various kinds of combo boxes and list boxes.

I've developed this ActiveX control in 2009 and did update it on a regular basis until 2016. Currently I have little interest to maintain this project any longer, but I think the code might be of some use to others.

# Before you make changes
If you make changes to the code and deploy the binary, keep in mind that ActiveX controls are COM components and therefore should stay binary compatible as long as you don't change the COM object's, i.e. the ActiveX control's public class name and GUIDs. Otherwise people using these components are likely to end up in the famous COM hell.

# Dependencies
You will need the [Microsoft Windows 10 SDK](https://developer.microsoft.com/en-us/windows/downloads/windows-10-sdk), ATL, and [WTL 10](https://sourceforge.net/projects/wtl/).

## Required Change to ATL
Some version of ATL have a bug in ```AtlIPersistPropertyBag_Load``` which causes crashes. In the file atlcom.h search for ```AtlIPersistPropertyBag_Load```. Inside the implementation of this function search for ```HRESULT hr = pPropBag->Read(pMap[i].szDesc, &var, pErrorLog);```. Make sure that there is this code before this line: ```var.pdispVal = NULL;```.