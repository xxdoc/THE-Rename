BEGIN { 
	printf "File=%s\n",ARGV[1] 
	temoin=0
	lastline=0
}

function Imprime()
{
	if (lastline != FNR)
	{
		printf "%d -> %s\n",FNR,$0
	}
	lastline=FNR;
}
/[:space]*Caption *=.*".*"/ { Imprime() }
/[:space]*ToolTipText *=.*".*"/ { Imprime() }
/[:space]*SimpleText *=.*".*"/ { Imprime() }
/[:space]*Title *=.*".*"/ { Imprime() }
/[:space]*Text *=.*".*"/ { Imprime() }
/[:space]*Title *=.*".*"/ { Imprime() }
/[:space]*TabCaption\([0-9]*\) *=.*".*"/ { Imprime() }
/Attribute VB_Name/ { temoin=1 }
/"([:alnum:]|[:blank:]|[:digit:]|[:punct:]|[:space])*"/ { if (temoin==1) 
							  printf "Temoin est vrai %d -> %s\n",FNR,$0 }
END { printf "-----------------------------------------------------------------------------\n" }
