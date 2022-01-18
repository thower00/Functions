let
    fnReadData = (ParameterTableName as text,ParameterName as text) as table =>

    let
        LoadFromTable = fnGetParameter(ParameterTableName,ParameterName,"LoadFromTable"),

	value = 
	if  LoadFromTable then 
    	    Excel.CurrentWorkbook(){[Name=fnGetParameter(ParameterTableName,ParameterName,"TableName")]}[Content]
	else if not LoadFromTable then
            Expression.Evaluate(fnGetParameter(ParameterTableName,ParameterName,"QueryName"),#shared)
        else
            "#FEL: fnReadData: LoadFromTable not defined"
    in
	value
in
    fnReadData