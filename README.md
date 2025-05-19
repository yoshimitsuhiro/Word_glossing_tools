A simple macro for Microsoft Word to set custom tab stops across a multi-tier interlinear morphemic gloss.

Simply import Tabulate_Glosses.bas and Tabulate_Glosses_Prompt.frm (make sure Tabulate_Glosses_Prompt.frx is in the same folder when importing) in Visual Basic for Applications, then highlight one of more lines of an interlinear gloss and run the macro.

Set the Indent and Interval between each element (both in mm) in the user prompt.

If Indent is set to "auto", the macro will calculate the width of the indent based on the first line of text (not including any automatic numbering).

If Interval is set to "auto", the macro will calculate the maximum interval between elements possible before elements begin to wrap over to another line. Set the maximum width (in mm) allowable in Max Interval. This setting is ignored if you input the Interval manually. 
