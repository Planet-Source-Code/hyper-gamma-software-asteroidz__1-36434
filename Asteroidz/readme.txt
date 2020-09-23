
Project Info...
 - Author:          Zach "Orion" Collette
 - Company:         Hyper-Gamma Software
 - Contact:         kingtheoden17@hotmail.com
 - Project Name:    "Asteroids"
 - Description:     Simple game using OOP
 - Test Platform:   Windows 98 SE
 - Processor:       P3 450MHz.
------------------------------------------------------------

Rules:

  Destroy all asteroids without running into them.  After all of the asteroids are destroyed, you advance to the next level.  It will get difficult quickly.  There is a penalty for running into an asteroid, and you lose health.  If you die, you lose 5,000 pts. and must restart the level with a different ship.  Lose all your ships, and the game is over.


+--------------+----------+-----------+-------------+
| Asteroid     |   Kill   | Collision | Health Lost |
+--------------+----------+-----------+-------------+
| Large (Red)  | 600 pts. | -300 pts. |     20      |
| Med (Brown)  | 450 pts. | -225 pts. |     15      |
| Small (Grey) | 300 pts. | -150 pts. |     10      |
| Tiny (Blue)  | 150 pts. | -125 pts. |     05      |
+--------------+----------+-----------+-------------+


+---------------------------------------+
|              Controls                 |
+---------------------------------------+
| Exit                - 'Escape'	|
| Pause               - 'F1'		|
| Toggle FpsCounter   - 'F2'		|
| Toggle Trails       - 'F3'		|
| Toggle Sound        - 'F4'		|
| Toggle Turbo        - 'F5'		|
| Turn Left           - 'Left'		|
| Turn Right          - 'Right'		|
| Accellerate         - 'Up'		|
| Deccellerate        - 'Down'		|
| Fire Ion Pulse      - 'Space'		|
| Fire Seeker Missile - 'Right Control' |
+---------------------------------------+

Description:

  This code makes heavy use of object oriented programming.  It should run smoothly enough.  It is not by any means optimized, it is structured for easy understanding, organization, etc.  It runs at about 30fps on my P3 450mhz machine.  The main game loop is in the 'Execute()' function in cls_Main.  I am sorry it is hard to follow since it isn't all commented.  All the files... sounds, bitmaps, etc are included in the resource file included in the project.  Please ask and me and I will send you the files.  To open it, goto AddIns->AddInManager and select resource editor, then check Loaded/Unloaded and press OK.  Please visit my wannabee site (http://www.geocities.com/hypergammasoftware) if you have the time.  Any questions/comments, please mail me at kingtheoden17@hotmail.com.  I will explain any aspect of this program for you as best as I can. :)


Features:
  OOP
  BitBlt
  Sound
  uses pure vb and windows api, no 3rd party controls or directx (making it more compatible)
  Collision Detection
  custom font (using fontmap)
  Custom Command Buttons (on frm_Splash, not using any controls, look at code...)
  bitmap classes (sprite support built in)
  enumerated input class
  Game physics
  Unlimited number of Asteroids, Shots, Effects, and keys to moniter (cls_Input).
  game effects
  seeking missiles
  polymorphism
  top ten high scores saved as random access file
  game config loaded from random access file
  loading files from resource file (.wav, .gif, string tables)
  Conver .gif format into .bmp formate (easier thatn you think!)
  changing screen resolution
  compartmentalized, most class modules can work independenly (encapsulation)
    

Stand Alone Classes (can work independantly in a different project):
  cls_Input
  cls_Res
  cls_Bitmap (works best with cls_Bitmaps)
  cls_Sound (also works best with cls_Sounds)


  Feel free to use this and distribute it as well, but as with all submissions, please do not alter in any way or take credit for this work.  Please do not use the graphics without my permision, I put a lot of work into them.  The sounds are free to use, though.  Thank You... Enjoy.
