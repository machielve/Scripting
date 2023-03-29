using Ridder.Client.UIScript;
using System.Drawing;
using Ridder.Common.Script;
using System;

public class RidderScript : ConditionalFormatScript
{
	public void Execute()
	{
	    //BITO-DE_3 regel 1	
        if (FindItem("TCR").Value.ToString() == "18" && 
		this.Item.Value.ToInteger()!=10000)
		{
			this.Item.Control.SetStrikeOut(true);
			this.Item.Control.SetBold(true);
			this.Item.Control.SetItalic(true);
			this.Item.Control.SetBackgroundColor(Color.Yellow);
		}

        //BITO-DE_3 regel 3
		if (FindItem("TCR").Value.ToString() == "20" && 
		FindItem("TP").Value.ToDouble()<0.261 &&
		FindItem("TP").Value.ToDouble()>0.259)
		{
			this.Item.Control.SetStrikeOut(true);
			this.Item.Control.SetBold(true);
			this.Item.Control.SetItalic(true);
			this.Item.Control.SetBackgroundColor(Color.Yellow);
		}




        //BRNZEEL_3 regel 1	
        if (FindItem("TCR").Value.ToString() == "18" && 
		this.Item.Value.ToInteger()!=10000)
		{
			this.Item.Control.SetStrikeOut(true);
			this.Item.Control.SetBold(true);
			this.Item.Control.SetItalic(true);
			this.Item.Control.SetBackgroundColor(Color.Yellow);
		}

        //BRNZEEL_3 regel 3
		if (FindItem("TCR").Value.ToString() == "20" && 
		FindItem("TP").Value.ToDouble()<0.451 &&
		FindItem("TP").Value.ToDouble()>0.449)
		{
			this.Item.Control.SetStrikeOut(true);
			this.Item.Control.SetBold(true);
			this.Item.Control.SetItalic(true);
			this.Item.Control.SetBackgroundColor(Color.Yellow);
		}



	}
}