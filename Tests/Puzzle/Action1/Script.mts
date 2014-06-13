r=inputbox ("enter no of rows")
c=(2*r)-1
m=(c-1)/2

ReDim a(r-1,c-1)
For i=0 to r-1
    For j=0 to c-1
        a(i,j)="*"
    Next
Next
'mid of all rows
For i=0 to r-1
  a(i,r-1)=1
Next


'rows manipulation
For i=1 to r-1
    For j=0 to m-1
                'If previous row has a value at next column position,increment it by 1
        If a(i-1,j+1)<>"*" Then
            val=a(i-1,j+1)+2 '2 instead of 1 because i am decrementing 'val' by 1 in next line
            For k=j to m-1
                 val=val-1
                 a(i,k)=val
                a(i,c-1-k)=val
            Next
        End If
    Next
Next

'print
For i=0 to r-1
    For j=0 to c-1
      txt= txt& a(i,j)
    Next
    txt=txt&vblf
Next
msgbox txt