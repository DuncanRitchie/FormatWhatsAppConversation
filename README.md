# Format WhatsApp Conversation
Console-app to format the text backup of a WhatsApp conversation into a Word document.
Pretty niche, and not the most user-friendly, but it does what I needed it to do.

## Input
The messaging app WhatsApp allows you to export a conversation to a _chat.txt file. It looks like this:

<pre>

[01/03/2019, 19:39:43] ‎You created group “Animals Talking”

[01/03/2019, 19:39:43] Animals Talking: ‎Messages and calls are end-to-end encrypted. No one outside of this chat, not even WhatsApp, can read or listen to them.

[01/03/2019, 19:40:10] Dog: Woof!

[01/03/2019, 14:44:50] Cat: Miaow, miaow, miaow!

[04/03/2019, 11:20:21] Rabbit: Appropriate rabbit noises

[04/03/2019, 11:33:57] Cat: Miaow miaow

</pre>

## How to use this app
Prerequisites:
* Windows,
* Visual Studio (just because I haven’t made a publish build), and
* Microsoft Word (because this app uses the Word interop API).

To use this app, you will need to: 
* clone this repo,
* move the text file into the root directory,
* build the Visual Studio solution, and 
* then run the program (eg by clicking “Start” in Visual Studio).

The app will immediately start processing the text, to eventually save a formatted-chat.docx Word document in your Documents folder.

## Output
Given the input above, formatted-chat.docx will look like this:

<pre>

Animals Talking

You created this group at 01/03/2019, 19:39:43

Dog — 01/03/2019, 19:40:10
Woof!

Cat — 01/03/2019, 14:44:50
Miaow, miaow, miaow!

Rabbit — 04/03/2019, 11:20:21
Appropriate rabbit noises

Cat — 04/03/2019, 11:33:57
Miaow miaow

</pre>

The group name “Animals Talking” is styled as a Title.

The author names and date-times are styled as Heading 2.

The other text is styled as Normal.

