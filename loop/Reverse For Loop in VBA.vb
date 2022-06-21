Dim intLoop As Integer
For intLoop = (Forms.Count - 1) To 0 Step -1
    DoCmd.Close acForm, Forms(intLoop).Name
Next intLoop