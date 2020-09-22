hi there

this project was created to push a new component into beta-testing. component's work name is "AniTrans" (animated and transparent). it's written in VB6 and tuned for VB6. at first glance it may look like just another animated GIF control, and, basically, it is so. but generally it pretending to be kinda "super easy to use" _animated_sprite_ control that makes it really easy creating applications with fairly complex 2D-animations without any additional DLLs nor stuff like DirectX.

there's two big things about it - transparent animations and pixel-counting collision testing. have you ever try to do graphics manipulations in your projects you may already know what a headache it may be. this control allows you to treat simple animations (taken from GIF89 files) just like the VB's native image controls. you drop the control on the form, load a GIF file into it - and that's it. it's playing. you manipulate it in standart way, eg. you move it by changing Left and Top properties.  it makes all dirty work for you - maintaining picture masks, drawing transparent images, keeping correct z-order positions etc. the control is windowless, so it doesn't eat unneccessary resources. however, that doesn't mean it's "lightweight" - it has to keep all frames of animation in memory unpacked so it takes a bit of memory if you know what I mean. 

control is able to stretch and flip animation frames. however, be aware that it slows redrawing speed a lot. also, it can sink mouse events only for "visible" parts of images.

it can use animation frames from another AniTrans control which has frames loaded (CloneFrames method). this means, for example, that you may have ten similiar animations on the form that share one animation source - thus saving resources.

the easy of use isn't costless - so don't expect having hundreds of fast animations at one screen running smooth. however, it should handle dozen of mid-size animations at 10-15 FPS each. of course, that depends on PC it's running on (more on videocard than CPU speed). 

the main form in this project (stress test form) should give you an idea of what could you expect from the control. it loads GIF files from GIFS\ folder and displays them all animated. use popup-menu (right-click on images). also, this form gives you idea of what is pixel-counting collision testing - drag any animation onto yellow smiling face and watch an indicator (rectangle) above the face. of course, step through the source code to get an idea how to use it.

second form shows how to create some  "meaningfull" animations using control's events and Z-order stuff. note that nice alligator walks "through" the static picture. examine the source to catch that simple trick (actually, there are two image controls representing static picture, one is on top of ZOrder and other is on bottom. walking alligator is between them in Z-order)

third form shows all techniques you need to create "behaviors" of sprites. dinamic loading and unloading, motions and collision testing. stupid aliens are bombing everything that moves. even themselves :) bombs are blowing as well as aliens :) targets have their lives indicated above them and once they "die" they resurrects again.
don't expect any AI in this sample - bombs are dropped randomly and targets just moves back and forth. it's not a game, just "simple sample".

and, finally, Form4, for "novices". it gives a glue about nuts'n'bolts - it draws each frame of animation as well as frame's masks separately. then, it combines them and - bingo - here it is, transparent picture

last word - since it's beta version, it has almost no docs you should have some "above average" VB experience  to get the glues just from poperties/methods names. if you ever tried to code games (or other graphics app) in VB you should get it pretty easy and appreciate how much dirty work does this control take from you.

just one note for quick start - use the "About" entry in the property browser of VB IDE to load GIF's. yes, that's nasty :) but quick and efficient ;) still no time for decorations like property pages etc.

any feedbacks are welcome at dmitrya@thewercs.com

/damian
