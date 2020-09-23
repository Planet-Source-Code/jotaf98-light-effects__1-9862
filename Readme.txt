		+---------------------------+
		| DRAWLIGHT FUNCTION README |
		+---------------------------+

		    +-- INTRODUCTION: --+

 "DrawLight" is a function that I wrote because I wanted to
add lightning effects to my game, if it was possible. Guess
what, it is -- and here's the proof! :)

 It might be too slow to do it in real time in games, but it
could be pre-calculated when loading a level or something.

 It draws a light in any picture, but it's faster than most
ways to achive this (because I used the circle's core
equation instead of the Circle function). Also, it's
highly customizable: you can change the coordinates where it
will appear, the radius, and, most important, the brightness
that will be added for the red, green and blue values.

 As a bonus, it draws the light like if it was growing, so
maybe there's no need for using a backbuffer ;)

 It should teach you something about using equations with
loops, messing with the RGB values of each pixel to get some
nice effects, and how basic algorythms for graphics work.

 There's only a small problem: if you use it twice in the
same image, and the lights "touch" each other, some weird
colors appear... if you know how to solve it, please e-mail
me.

		    +-- INSTRUCTIONS: --+

 Simply add the .bas file to your project to use it.
 All you need to do in order to activate the effect is to
call this function.

			+-- USAGE: --+

 DrawLight(Target As Long, X As Single, Y As Single, RedB
As Byte, GreenB As Byte, BlueB As Byte, Radius As Integer,
NumberOfSteps As Single)

		      +-- ARGUMENTS: --+

+- Target:
 This is the hDC where the image will be drawn.
(You'll find this as a property of forms and picture
boxes.
 Just know that it refers to the image used in these
controls.)

+- X:
 The X coordinates of the place where the light will be
drawn (referring to the center).

+- Y:
 The Y coordinates of the place where the light will be
drawn (referring to the center).

+- RedB
 This is the brightness applied to the Red value. Be
aware that high numbers will make the light totally
white!

+- GreenB
 The brightness applied to the Green value.

+- BlueB
 The brightness applied to the Blue value.

+- Radius:
 The radius of the light.

+- NumberOfSteps
 The light is drawn in various steps (circles that
apply different brightnesses to the background). The
higher the steps, the smoother the light will appear,
but too many steps make the process slower.

		       +--EXAMPLE: --+

 Check the sample project included for a working demo.
 But the .exe runs much faster than Visual Basic's Run
Mode!

		+--------------------------+
		|  Made by Jotaf98         |
		|  (João F. S. Henriques)  |
		|                          |
		|  E-mail me at            |
		|  jotaf98@hotmail.com     |
		+--------------------------+