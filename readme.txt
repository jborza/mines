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

TODO:
Colors per different number - but cannot change it on a CommandButton

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