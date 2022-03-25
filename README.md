## This was a 1 day project for making icons for movies
It's only here because, someone might be able to extract some useful code for a similiar project.
Nowadays I rarely code in C# and this was created really fast.

## Process
Gets all movies from a specified folder, then sequentially jumps through the current movie x times, takes a screenshot, then creates an ico file from the png. 
which then gets saved to the specified folder.
After that, it can copy the contents of the folders to a new directory (path) and set it's icon to the choosen one.

### Basic usage
 - Set the ```sourceDir```, ```iconPath```, ```newDirectory``` variable values
 - Run the project, then click browse
 - Choose the icon best fit for the movie when it finishes from iconPath/x and rename it to ```icon```
 - After that delete all other icons from the directory
 - Then click "Take to new with icons" labeled button
 - When it finishes it should list the moved folders

### Issues
There are a lot. Most methods are not defined well and very specifically were created and used for this project. 
Biggest problem is with the resolution, which can't be changed. I didn't want to bother calculating the user resolution in perspective to the created images. It can be done though, if you're determined enough and it might not even be a long time.