let
    fnTruncateText = (TextString as text, MaxLength as number) as text => 
    let
        TruncatedText =  
        if Text.Length(TextString) >= MaxLength then
            Text.Start(TextString,MaxLength-2)&".."
        else
            TextString

    in
        TruncatedText
in
    fnTruncateText