/* -----------------------------------------------------------------
   Table of Content scripts
   http://www.mytreedb.com/treedbnotes_pro.html
----------------------------------------------------------------- */

function Expander(ElementID)
{
	var MyDiv = window.document.getElementById('DV'+ElementID);
	var MyImg = window.document.getElementById('IM'+ElementID);
	var MyCnt = window.document.getElementById('CN'+ElementID);
	var imgSrc = '';
	var cntSrc = '';
  if(MyImg != null) { imgSrc = MyImg.src; }
	if(MyCnt != null) { cntSrc = MyCnt.className; }
	if ( MyDiv != null )
	{
		if ( MyDiv.style.display == "none" )
		{
			MyDiv.style.display = "";
			imgSrc = imgSrc.replace("plus", "minus");
			cntSrc = cntSrc.replace("plus", "minus");
		} else {
			MyDiv.style.display = "none";
			imgSrc = imgSrc.replace("minus", "plus");
			cntSrc = cntSrc.replace("minus", "plus");
		}
		if ( MyImg != null && imgSrc != "" )
		{
		  MyImg.src  = imgSrc;
		  MyImg.hsrc = imgSrc;
		}
		if ( MyCnt != null && cntSrc != "" )
		{ MyCnt.className  = cntSrc; }
	}
}
