using Ridder.Client.UIScript;
using System.Drawing;
using Ridder.Common.Script;
using System;

public class RidderScript : ConditionalFormatScript
{
	public void Execute()
	{
		if (FindItem("TCR").Value.ToString()

		== "18" && //BITO-DE_3 regel 1
		this.Item.Value.ToInteger()!=10000)

		{
			this.Item.Control.SetStrikeOut(true);
			this.Item.Control.SetBold(true);
			this.Item.Control.SetItalic(true);
			this.Item.Control.SetBackgroundColor(Color.Yellow);

		}

		if (FindItem("TCR").Value.ToString()

		== "20" && //BITO-DE_3 regel 3
		FindItem("TP").Value.ToDouble()<0.261 &&
		FindItem("TP").Value.ToDouble()>0.259)

		{
			this.Item.Control.SetStrikeOut(true);
			this.Item.Control.SetBold(true);
			this.Item.Control.SetItalic(true);
			this.Item.Control.SetBackgroundColor(Color.Yellow);

		}

	}
}