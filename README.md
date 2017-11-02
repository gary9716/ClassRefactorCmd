# Visual Studio Extension Project - Class Refactor Utility
## Introduction
It is a visual studio extension project that demonstrate how to rename a class name programmatically in script file written in csharp for example.
It utilize the refactoring function provided in Visual Studio.

## How to use
1. download this repository
2. unzip it and open it with visual studio(I have tested in VS2015 and it should also work in VS2017 and maybe in VS2013)
3. after the project is loaded, press the start button and VS will start building this project
4. if the build succeed, another Visual Studio instance would show up(it's an experimental instance).
5. press the "Tools" tab on menu bar and press the menu item with the text "rename class".
As the menu item being pressed, "MenuItemCallback" function in MenuCmd.cs will be called.
The renaming logic is implemented in this function and I also wrote some comments there so it should be comprehensible.
