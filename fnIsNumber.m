let
   fnIsNumber = (lookupValue as any) as logical =>
   let
        result = try Number.From(lookupValue) otherwise false,
        resultType = if result <> false then true else false
   in
        resultType
in
   fnIsNumber