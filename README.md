# 4d-plugin-vbs-outlook

Experimental

Since Office 2013/2016, Outlook automation has become very difficult. Click-To-Run (a.k.a. C2R) deployment means Outlook no longer exposes interfaces such as [``IConverterSession``](https://msdn.microsoft.com/en-us/library/office/ff960231.aspx) to COM. This one was useful to convert MAPI (``msg``) to MIME (``eml``). Microsoft has decided not to exposes interoperability classes in the common namespace but rather insulate them in their virtual namespace ([bubble](https://blogs.msdn.microsoft.com/stephen_griffin/2014/04/21/outlook-2013-click-to-run-and-com-interfaces/)).
