-----------------------------------------------------

Version 3
December 28, 2000

I added the ability to save and load the region data to a file.  This drops the Form1 create time mentioned below from ~1.2 seconds using the brute-force method to ~.007 of a second.  You do this by setting the first optional parameter in the ShapeMe method to False one time during development time to generate a region data file.  Then set it to True to load the data file after that.  See the example on Form1.  I found the code to load and save region data on vbaccelerator.com.

-----------------------------------------------------

Version 2
November 3, 2000

Summary for use:
* Set the Picture (and DownPicture) property in code or at design-time.  For a Form or PictureBox, make the control just big enough to expose the graphic.  For a CommandButton, CheckBox or OptionButton make it slightly larger (see notes below).

* Set several other properties at design-time (see notes in the code in clsTransForm).

* Instantiate the class to expose its methods (see code for an example).

* Use ShapeMe to shape the object.
	*Form: use this in Form_Load
	*PictureBox: can use this in Form_Load or Form_Activate
	*CommandButton, CheckBox, OptionButton: must use this in Form_Activate.  Add code so this code only runs once (see example), otherwise weird things happen.  Add a DoEvents after shaping all Buttons that are on one form so they draw correctly.  UpdateWindow API didn't get rid of this problem (?).

* Use DragForm in Form_MouseDown if you want your users to be able to move the form.
-----------------------------------------------------
Other notes:
This class allows you to have a Form, PictureBox, CommandButton, CheckBox or OptionButton shaped exactly like the image you assign to its PICTURE property.  A specified background color, in this case white, that exists in the picture will be made transparent (and non-existent since you can click on objects behind it).  Other code allows you to drag the form around since it doesn't have a title bar.

Originally based on code by Chris Yates (Automatic Form Shaper) from www.planetsourcecode.com, but modified completely and only shares about 5 lines of code with the original.  My version runs much faster and will work on more objects than just a Form.  Should work with any object that has an hDC, but I have restricted it for my own use.  I also replaced the code from the original to move the form because it didn't work (at least not in NT).  My test form image (the one in Form1) took ~27.5 seconds to create on a PIII-650 with Chris's code.  I cut it down to ~1.6 seconds in version 1 and then ~1.2 seconds in version 2. 

The time savings were due to: cutting the CombineRgn calls down from many, many thousands (one for every single point) to a few hundred in many cases (by searching for adjacent points and combining them into one line)(cut 25.4 seconds); I removed the HorizontalScan option that I added in version 1 because after much testing it didn't appear that there was much time saved by a vertical vs. horizontal scan on all the pictures I tried(cut .1 secs); I completely rearranged some code from version 1 so TypeOf was only called once and not once for each loop (cut .4 secs).  Other misc. cut .4 secs.

If you can cut the CombineRgn calls down even more, great, but you won't be able to except for very specific graphics.  The only way to do that is to write more code to search for exact rectangles (or circles if you use CreateEllipticRgn, or specific weird shapes if you use CreatePolygonRgn).  Statistically most pictures will not have them, so you waste lots of CPU power (and time) looking for them all the time.  If you just combine adjacent lines into rectangles you don't save any CombineRgn calls:

Example:

***************
***************X*****
***************X*****X**

Here you could cut lines 2 and 3 where I put the Xs, and make one call for each rectangle, and one for the remaining line (three calls total).  However, if you had just stuck with making a call for each line, you still would only have three calls.  Don't waste your time writing extra code to do this, I almost did until I worked this logic out.

Make sure you make the Form or PictureBox size just big enough to expose the graphic, unless you intend on exposing an edge or two (if the background color does not match the transparent color; if they match the exposed form edges will disappear too, but it wastes lots of CPU power and time).  You may want to do that to put controls on the exposed edges?  You must make a CommandButton, CheckBox or OptionButton slightly larger than the graphic because the edges get chopped off to hide the button's border (which doesn't disappear even if you tell VB to make it "flat" - a continuing bug they may never fix).  Note that VB or Windows adds highlighting to a button when pressed.  This may be the system Button Highlight color, but I haven't tested.  You can get around this by making your button's border the same color as the highlighting, or add subclassing, as described below.

I have included the code in a class module.  I dropped the standard module from version 1 because you could add subclassing for the Form to the class module so you can get rid of the DragForm method and let the class handle it automatically.  You might also add subclassing to the Buttons and let the class handle how they are created to get rid of the kludgy DoEvents and other code necessary to make them draw right.  That would also get rid of the highlighting when you press a button.  Thus only a class here in version 2, for ease with future expansion.

I also cleaned up my coding a little in this version...I was rushed before.  Now I start the program with a Sub Main (good programming practice all the time) and changed the error notice to a real err.raise instead of a messagebox (which is very, very bad practice, but worked while I was testing it).

I have only tested this code in NT4, SP4.

Email me at dgaede@home.com with comments or questions.

Hope this code helps you out!

Doug Gaede