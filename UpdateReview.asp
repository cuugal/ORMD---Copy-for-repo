<%
   if NoSaveBeforeSWMS <> "nosave" then
        ' This segment sets the Review Date (dtReview) in tblQORA.
        ' Review date is the nearest future (i.e. not predating today) date for the appropriate RA in tblRiskControls
        ' or the date a year ahead of creation/edit, whichever is soonest/not null.
        set rsUpdateReview = Server.CreateObject("ADODB.Connection")
        rsUpdateReview.open constr
  	    ' MS Access has no support for aggregation within updates, so we need to do this in two parts.
  	    ' First, get the new review date.
  	    strSQL2 = "SELECT iif(strAssessRisk = 'L',  DateAdd('yyyy',3,date()),  DateAdd('yyyy',1,date() )) as proposedDte, "_
				    &"IIF(IsNull( min(rc.dtProposed) ), proposedDte, IIF(min(rc.dtProposed) < proposedDte ,min(rc.dtProposed) ,proposedDte )) as NewDate "_
				    &"from tblQORA "_
				    &"left outer join tblRiskControls rc on rc.numQORAID = tblQORA.numQORAID "_
				    &"where rc.dtProposed >date()  or IsNull(rc.dtProposed) "_
				    &"group by tblQora.numQORAID, tblQora.strAssessRisk "_
				    &" having tblQora.numQORAID = "&testval
				
	    set rsShowReview = Server.CreateObject("ADODB.Recordset")
	    rsShowReview.Open strSQL2, rsUpdateReview, 3, 3 
	    dim dtNew	
	    if not rsShowReview.EOF then 		
		    dtNew = rsShowReview("NewDate")
	
		    'Now, run the update
		    strSql3 = "Update tblQORA set dtReview = '"&dtNew&"' where numQORAID ="&testval
	
  
  		    set rsAdd = Server.CreateObject("ADODB.Recordset")

	    'Executing the SQL this way ensures we have an exclusive database lock.
	    'The access database is inherently multithreaded, and there exists the case where
	    'a read immediately after an update will fail to yeild the changes, simply because the update was in a different 
	    'thread that took longer.
	    ' Here we lock the db, preventing any reads until we have finished with our changes.
   	 	    rsUpdateReview.BeginTrans
  		    rsUpdateReview.Execute strSQL3
  		    rsUpdateReview.commitTrans
  	    end if
    end if
  %>
  