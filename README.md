# MAGI
### Madcap Automated Graphic Implementation - uses API access with Monday and Sharepoint to automate graphic amends in Flare

This project is pretty niche! It is used to automate graphic embedding in MadCap Flare, but in line with specific systems. Two other apps are used - Monday.com and Microsoft's SharePoint.

 - **Monday** acts as a task manager/sprint board, mainly for businesses. In this case, it is relevant for managing graphic progress by adding items to the relevant board.
 - **SharePoint** acts as a file manager. In this case, the finished graphics are uploaded to the relevant locations in a SharePoint file system.
 - **MadCap Flare** is essentially a front end dev framework, allowing word doc imports and easy HTML/CSS editing. In this case, graphics are embedded to Flare to bring together a finished product.

**At its current stage, MAGI is not easily redirected to different endpoints without a knowledge of Python.**

## What does MAGI do?

MAGI contacts Monday's API to scrape a list of graphic names to embed in your project. It will fetch them from a board selected by you, automatically saving the board's ID to a config for future use.
Once you select a specific Monday board, it will scrape the name of _all items_ from that board. It will then attempt to cross-reference these with SharePoint.

MAGI then contacts SharePoint's API to match all graphic names from Monday and download the file related to each name. It is currently downloaded to a specific folder in your Flare project
that may not exist or may be irrelevant to your project - this can only be changed by editing the code's main body.

Finally, it will scan HTML documents in a folder withing your project, selected by you, in order to embed them. For this, the file `baseSnippet.html` can be edited to change what
specifically is embedded in the graphic name's place. Currently, it is a figure tag with an image and a caption. It then creates a duplicate HTML file with all the successful embeds for you to manually
check. This is done mainly for sanity as instantly overwriting files could break or corrupt them.

## Where do I add my API details?

MAGI uses eel, a Python framework, to generate a GUI for ease of use. In the window that appears, you can enter your details:

 - SharePoint username
 - SharePoint password
 - Monday API key

Following this, they are saved to separate config JSONs. These are unpacked each time you run the program so that you never have to re-enter them.

## What else?

MAGI can check the folder where it saves your graphics and generate a report of how many have been successfully embedded in your project. It will create a `.txt` report of
each graphic that for one reason or another is **not** in the project. This is split into those not found at all, and those which failed the embed process. The latter will be quoted
with the graphic name, the document they are found in and the line they are on.

---

**Realistically, MAGI is not versatile enough to be used by anyone that doesn't have the same specific setup as the author. It is intended as a demonstration of Python application writing.
For this reason, the README does not go into detail.**
