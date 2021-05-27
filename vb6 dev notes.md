Created new project
- timers work as in .net

Colorful User control acts as a special control - we can use it for mines and stuff

it gets painted

## Terminology

Minefield, square

## Project structure

### Mine - User Control
It will have a property that is the state
It will have one button that will display the state

### Field - User Control
A grid of mines
We'll spawn mine controls within the field according to its configuration

Defined the Point type in Module

Global variables for the settings such as Rows, Columns, Mines

Form - displays timer, mine counter

## Generating the minefield
We also need to check if the mines don't overlap - TODO



## How to add custom controls to the form?
Close the control designer to make it available in the form designer

## Return inside a function
We use `Exit Function`

## Woes
Renaming a control doesn't rename all references in the code. 

Could not add square control on its own - had to uncheck Remove information about unused ActiveX controls

Could not figure out a way to trigger flood fill from the child control, the syntax is actually:

Definitely set "Save changes" option in VB6 settings->Environment->When program starts

The game sometimes crashes with access violation

```vb
Call ParentControls(0).Floodfill
```

mines at 3x3 with 3 mines: 3,2 2,1, 1,3
```
 X 
  X
X  
```

should have numbers as:

```
1X2
23X
X21
```

### mines at 4x4:


```
0111
02X2
02X2
0111
```

Mine indices: [6, 10]

More notes:

Flood fill:
https://en.wikipedia.org/wiki/Flood_fill

```
The earliest-known, implicitly stack-based, recursive, four-way flood-fill implementation goes as follows:[4][5][6][7]

Flood-fill (node):
 1. If node is not Inside return.
 2. Set the node
 3. Perform Flood-fill one step to the south of node.
 4. Perform Flood-fill one step to the north of node
 5. Perform Flood-fill one step to the west of node
 6. Perform Flood-fill one step to the east of node
 7. Return.

```

However, for Minesweeper we should do diagonal minefill.

I wanted colors per different number on the revealed mine neighbors - but cannot change it on a CommandButton. I had to eventually introduce a new control for this that consisted of a Label and both the label and the user control were clickable.

Very nasty bug with collections - when I was using key as variant, the col.Item(key) 
returned True for key=1, when there was **one** item in the collection.

Public Function Exists(ByVal col As Collection, ByVal key As String) As Boolean
    On Error GoTo DoesntExist
    col.Item (key)
    Exists = True
    Exit Function
DoesntExist:
    Exists = False
End Function

## Debugging

Immediate Window works as in newer Visual Studios:

```
? MineLookup.Count
 3 
? TypeName (MineLookup)
Collection
```

Do we color the discovered 

Color:
#C6C3C6 -> &HC6C3C6&
&H00E0E0E0&

Could not figure out how to remove controls from the item array, it had to be invisible first, then could remove it with `Unload`

### Adding menu and menu handlers

We can invoke the menu editor with Ctrl+E, 

### Discovering syntax errors
File->Make tries to compile the code to native code and also stops and notifies the developer on a compile errror.