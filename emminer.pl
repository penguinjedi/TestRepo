#===============================================================================
# program emminer.pl                     
#
# Copyright 2005-2006 BMC Software, Inc. as an unpublished work. All rights reserved
#
# Author: Terry Cannon 
#         (with special thanks to all who continue to contributed ideas/corrections)
#
# Warranty: No warranty is expressed or implyed with this routine.  It is
#           not a product and does not have support for it.  If you use
#           it, you must answer any issues you may have with it.
#
# Purpose: environmental scans (primarily data sumation) on Control-M EM DB
#
# Tested on EM series 6** Control-M Enterprise Manager 
#
# Version number: Jan, 2006
#
# Syntax:   emminer.pl		{invokes as interactive)
#           emminer.pl -d	(turns on debugging)
#           emminer.pl -silent  (uses previous input values and runs with no prompts)
#           
#           or with a fully qualified path to where perl is installed, like:
#
#           c:\program files\bmc software\control-m em\default\bmcperl\perl emminer.pl
#           c:\program files\bmc software\control-m em\default\bmcperl\perl emminer.pl -silent
#
# Jan, 2007 updates
#	     - (TH) changed "quotes" used in the select of USERGROUP (worked for
#              sybase but needed single quotes for Oracle select, caused
#              invalid select when run against Oracle 
#            - (TH) Added SILENT option which will take previous input values and
#	       run without being prompted (in case you want to run as a batch job) 
#	     - (TH) Added the notification if your daily usage exceeds a Maximum number you give.
#	       This is in case you know what your maximum number of daily task are
#  	       and you would like the routine to inform you if you exceed that number.
#	     - (TLC) Deactivated the Sybase/MSDE query for DB Parms (sp_configure) as it has
#              not really been useful 
#	     - (TLC) Again I have turned off "nolock" in the selects against Oracle DBs.  I think
#              the issue is that the value nolock appears in a different place for oracle
#              selects than for sybase selects.  Someone even passed that info to me but I
#              have misplaced it for now.  When I confirm this I can make a quick adjustment
#              so it is in the correct place in the select statement for Oracle queries also.
#	       It is still in place for non Oracle DB queries.
# 	     - (TLC) Added what version of EM was being used on the Misc sheet
#
# Nov, 2006 updates follow
#	     - (PS) Various overall suggestions
#            - removed the scftp subroutine
#            - removed the ecsinstall subroutine
#            - removed the selfping subroutine
#            - removed the osdetail subroutine
#            - removed the getenv subtoutine
#            - removed the getreg subroutine
#	     - removed the logspace subroutine
#            - removed duplicate "command strings" for ctmfw
#            - various small code changes for overall cleanup
#            - added option to run against MSDE db (via osql)
#            - added additional debug statements which are activated by
#              invoking with "emminer -d"
#            - moved the updconfig (new routine) to execute earlier so that
#              user input is captured and updated at the start of the routine
#              instead of at the end (so if it failed it was lost)
#            - activated the "with nolock" code (had been turned off before)
#            - added a testdb routine to validate access to the db
#	     - added queries based on
#		   - Days of the Month string
#		   - Days of the Week string
#		   - Showing which security group each user belongs to
#		   - Doc lib usage
#		   - Override lib usage
#
#
# future ideas:
#	     - add back in the global prefix to/from dc.  Had an issue with Oracle
#              and the substr function.  Need to substring because of field size
#            - need to complete testdb routine for oracle to verify DB availability
#===============================================================================

use Getopt::Long;				# needed to accept command line parms

GetOptions( "d"      => \$debug,		# "emminer -d" turns on debugging
            "silent" => \$silent); 		# "emminer -Silent" turns on silent running 

use Win32::OLE qw(in with);			# for excel access
use Win32::OLE::Const 'Microsoft Excel';	# for excel access
$Win32::OLE::Warn = 3;                          # die on errors... for excel access

#$user_help="NO";				# set this variable to no to avoid extra msgs
$user_help="";					

#======================     MAIN function     ============================

print ("emminer.pl the data miner is collecting needed info ... \n\n");

&getconfig();					# access previous saved values for the upcoming prompts 						

#--------------------------------------------------------------------------
# test to see if we are running "silent" so that we do not prompt for input
#--------------------------------------------------------------------------

if ($silent)
    {
    	print "----- Silent run, all parms will be used from previous run\n";
    	goto Haveparms;
    }

&getuser_input();				# get the id, password, db type and server from the user
&updconfig();					# save user input (except password) for later runs

Haveparms:					# at this point we have all the input parameters (either from
						# the user interactively or from the values saved in the config file
						# from previous runs

&initvars();					# set up some initial variables
&startexcel();					# initiate an excel session in background if not already running  
&testdb;					# this routine verifies access to the db
&dbqueries;					# run a series of selects against the EM db  
&override_colwidth;				# set any specific column widths you want.
&wrapup;					# final details
&cleanup();					# close excel if needed and cleanup temp files
      
#======================     end of main function ===========================





#======================     dbqueries function   ===========================
sub dbqueries ()
{
  if ($debug) {print "--- debug dbqueries routine\n";}
  print "   --> Now mining the EM DB for job info\n\n";
#  print scout "\n\n---> Now running a series of queries against your EM DB\n\n";	
  
  open (temp,">c:\\temp\\emminer.report.01");
  print temp "$mysqlpre1";			# sets some sql environment (pagesize, tab off, ...)
  print temp "$mysqlpre2";  
   
  $querytot=43;					# current total of queries 
 
#deactivated TLC Jan, 2007  if ($dbtype eq "O") {$querytot=$querytot-1;}		# current total of queries is 1 less for Oracle (no DB parms)
  
  print "\n         Exploring definitions\n\n";	# print headings to the screen so user can see whats occuring
  print "    query#          description\n";
  print "    ------     -----------------------------------\n\n";
  
#deactivated TLC Jan, 2007   if (($dbtype eq "M") || ($dbtype eq "S") || ($dbtype eq "E"))  	# this query is done only for non-Oracle DB right now
#deactivated TLC Jan, 2007       {
#deactivated TLC Jan, 2007       	print "    $querycnt/$querytot       DB settings\n";	# create sql to capture DB parms from Sybase
#deactivated TLC Jan, 2007   	print temp "go\nsp_configure \ngo\n"; 
#deactivated TLC Jan, 2007  	dosql();				# call generic SQL runner subroutine
#deactivated TLC Jan, 2007         $current_sheet="DB Parms";
#deactivated TLC Jan, 2007         $putmethod="sep";			# used while changing from my substring parcing to a separator type parcing of columns
#deactivated TLC Jan, 2007   	putsheet(); 	 			# call generic parcing routine to put values into excel cells
#deactivated TLC Jan, 2007   	$querycnt=$querycnt+1;			# increment the querycnt (for user status of how done they are)
#deactivated TLC Jan, 2007       }
     
  print "    $querycnt/$querytot       Doclib\n";	
  print temp "select count(*) $mycountq1   #jobs per Doc lib       $mycountq2,$sep,DOC_LIB from DEF_JOB $nolock GROUP BY DOC_LIB ORDER BY DOC_LIB $go";
  $querycnt=$querycnt+1;
  dosql();
  $current_sheet="doclib";
  putsheet();

  print "    $querycnt/$querytot       Memlib\n";	
  print temp "select count(*) $mycountq1   #jobs per Mem lib       $mycountq2,$sep,MEM_LIB from DEF_JOB $nolock GROUP BY MEM_LIB ORDER BY MEM_LIB $go";
  $querycnt=$querycnt+1;
  dosql();
  $current_sheet="memlib";
  putsheet();
  
  print "    $querycnt/$querytot       Override Library\n";	
  print temp "select count(*) $mycountq1   #jobs per Over lib    $mycountq2,$sep,OVER_LIB from DEF_JOB $nolock GROUP BY OVER_LIB ORDER BY OVER_LIB $go";
  $querycnt=$querycnt+1;
  dosql();
  $current_sheet="overlib";
  putsheet();
  
  
    print "    $querycnt/$querytot       SNMP settings\n";	
    
  print temp "select $mysubstr(PNAME,1,80) $mycountq1 PName$mycountq2,$sep,$mysubstr(PVALUE,1,80) $mycountq1 Value$mycountq2 from PARAMS $nolock ";
  print temp " where PNAME=$myquote\MaxOldDay$myquote or PNAME=$myquote\SendAlarmToScript$myquote";
  print temp " or PNAME=$myquote\SendSnmp$myquote or PNAME like $myquote\Snmp$mypat$myquote  \n";
  $querycnt=$querycnt+1;
  dosql();
  $current_sheet="SNMP";
  putsheet();
        
  print "    $querycnt/$querytot       Components \n";
    
  print temp "select CURRENT_STATE $mycountq1     Current             $mycountq2,$sep,DESIRED_STATE $mycountq1 Desired $mycountq2,$sep1,$mysubstr(PROCESS_NAME,1,25) $mycountq1 Process$mycountq2,$sep1,$mysubstr(MACHINE_NAME,1,25) $mycountq1 Machine$mycountq2,$sep1,$mysubstr(PROCESS_COMMAND,1,80) $mycountq1 Command$mycountq2,$sep1,$mysubstr(ADDITIONAL_PARAMS,1,25) $mycountq1 Additional parms$mycountq2,$sep1,MACHINE_TYPE from CONFREG $nolock $go";
  $querycnt=$querycnt+1;
  dosql();
  $current_sheet="Components";
  putsheet();
 
  print "    $querycnt/$querytot       Tables by User Daily\n";

  print temp "select count(*) $mycountq1  #Sched tables per Daily$mycountq2,$sep,DATA_CENTER,$sep1,USER_DAILY from DEF_TABLES $nolock GROUP BY DATA_CENTER,USER_DAILY ORDER BY DATA_CENTER,USER_DAILY ";
  $querycnt=$querycnt+1;
  dosql();
  $current_sheet="Tbls per User Daily";
  putsheet();  
 
 
  print "    $querycnt/$querytot       Owners\n";

  print temp "select count(*) $mycountq1   #Jobs per Owner       $mycountq2,$sep,OWNER from DEF_JOB $nolock GROUP BY OWNER ORDER BY OWNER  $go";
  $querycnt=$querycnt+1;
  dosql();
  $current_sheet="Owner";
  putsheet();

  print "    $querycnt/$querytot       Authors\n";

  print temp "select count(*) $mycountq1   #Jobs per Author      $mycountq2,$sep,AUTHOR from DEF_JOB $nolock GROUP BY AUTHOR ORDER BY AUTHOR $go";
  $querycnt=$querycnt+1;
  dosql();
  $current_sheet="Author";
  putsheet();
 
  print "    $querycnt/$querytot       Intervals\n";
  $cyclic=1;
  $cyclic="$myquote$cyclic$myquote"; 
  print temp "select count(*) $mycountq1   #Jobs per Interval    $mycountq2,$sep,INTERVAL from DEF_JOB $nolock where CYCLIC=$cyclic GROUP BY INTERVAL ORDER BY INTERVAL $go";
  $querycnt=$querycnt+1;
  dosql();
  $current_sheet="Interval";
  putsheet();


  print "    $querycnt/$querytot       Priority\n";

  print temp "select count(*) $mycountq1  # of Jobs  $mycountq2,$sep,PRIORITY from DEF_JOB $nolock group by PRIORITY order by PRIORITY $go";
  $querycnt=$querycnt+1;
  dosql();
  $current_sheet="Priority";
  putsheet();

  print "    $querycnt/$querytot       Time Zones\n";

  print temp "select count(*) $mycountq1  # of Jobs  $mycountq2,$sep,TIME_ZONE from DEF_JOB $nolock group by TIME_ZONE order by TIME_ZONE  $go";
  $querycnt=$querycnt+1;
  dosql();
  $current_sheet="TimeZone";
  putsheet();

 
  print "    $querycnt/$querytot       Max Waits\n";
                    
  print temp "select count(*) $mycountq1   #by Max Wait          $mycountq2,$sep,MAX_WAIT from DEF_JOB $nolock GROUP BY MAX_WAIT ORDER BY MAX_WAIT   ";
  $querycnt=$querycnt+1;
  dosql();
  $current_sheet="Max Wait";
  putsheet();

  print "    $querycnt/$querytot      From times\n";
                              
  print temp "select count(*) $mycountq1   #Jobs by From Time    $mycountq2,$sep,FROM_TIME from DEF_JOB $nolock GROUP BY FROM_TIME ORDER BY FROM_TIME    ";
  $querycnt=$querycnt+1;
  dosql();
  $current_sheet="FromTime";
  putsheet();

  print "    $querycnt/$querytot      To times\n";
                                
  print temp "select count(*) $mycountq1   #Jobs by To Time      $mycountq2,$sep,TO_TIME from DEF_JOB $nolock GROUP BY TO_TIME ORDER BY TO_TIME   ";
  $querycnt=$querycnt+1;
  dosql();
  $current_sheet="ToTime";
  putsheet();
 
  print "    $querycnt/$querytot      Global prefixes\n";
  $fdc="FROM_DC";
  $fdc="$myprintq$fdc$myprintq";
  $tdc="TO_DC";
  $tdc="$myprintq$tdc$myprintq";
  $seps=$sep1; 
#have commented out the original query for prefixes which also list to and from data centers
#this query generated sql errors against the latest sybase versions.  it appears that you cannot
#execute a substring against a field which has an underscore in it like TO_DC or FROM_DC.  as a result
#this sheet was misalligned greatly.  So for now, it only shows the global prefixes.             
#  print temp "select $mysubstr(PREFIX,1,25) $mycountq1    Global prefixes      $mycountq2,$sep,FROM_DC $mycountq1 From Data Center$mycountq2,$seps,TO_DC $mycountq1 To Data Center$mycountq2 from GLOBAL_COND $nolock ORDER BY PREFIX $go";
  print temp "select $mysubstr(PREFIX,1,25) $mycountq1    Global prefixes      $mycountq2 from GLOBAL_COND $nolock ORDER BY PREFIX   $go";
  $querycnt=$querycnt+1;
  dosql();

  $current_sheet="Globals";
  putsheet();
 
  print "               --> identify jobs with Global In and Out Conditions\n";

  system "echo --------------- > c:\\temp\\emminer.report.01.out.inout";
  system "copy c:\\temp\\emminer.report.01.out c:\\temp\\emminer.report.01.out.gps >c:\\temp\\emminer.report.01.out.copy.txt";
 
  open (resultsin,"<c:\\temp\\emminer.report.01.out");
  while (<resultsin>)
    {
    	$l1=length($_);
    	if ($l1 > 2)
    	   {
    	     $gp_count=$gp_count+1;
    	   }
    }
  close resultsin;
  $gp_count=$gp_count-2;	# subtract 2 lines for the headings that SQL produced
  $gp_upd_interval=$gp_count/10;
  print "               --> total of $gp_count global prefixes to process, showing % completion on next line\n";
  print "                ";  
 
  open (gps,"<c:\\temp\\emminer.report.01.out.gps");    
  $current_sheet="Jobs w Globals";
  $gpprogress=0;				# counter to indicate activity

#-------------------------------------------------------------------------------
# for each global prefix, find all Jobs with either IN's or OUT's that are global
#-------------------------------------------------------------------------------

  while (<gps>)
      {

     	chop;
     	$glp=substr($_,1,25);
     	$glp =~ s/^\s+//;$glp =~ s/\s+$//;        #remove leading & trailing blanks
        $i1=index($_,"prefix");
        $i2=index($_,"-----");
        $l1=length($_);
        if (($i1 > -1) || ($i2 > -1) || ($l1 < 2))	# skip headings, blank lines, ...
           {
           	goto nextgps;
           }
        $gpprogress=$gpprogress+1;
 
 	if (($gpprogress == $gp_upd_interval) || ($gpprogress > $gp_upd_interval)) 
     	
     	   {
     	     $tot_gp_done=$tot_gp_done+$gpprogress;
     	     if ($gp_count < 1) {$tot_perc_done=100;}		# had a divide by zero early in testing when there were no globals
     	     else {$tot_perc_done=$tot_gp_done/$gp_count*100;}
             print "---";
     	     printf ("%3d",$tot_perc_done);
      	     print "$percent";
     	     $gpprogress=0;
     	   }	# show activity by update user on status	       
	                  
gins:   $cond="$glp";
        $cond="$myquote$cond$mypat$myquote";
        print temp "select DATA_CENTER,$sep,SCHED_TABLE,$sep1,JOB_NAME $mycountq1 Jobname $mycountq2,$sep1,$mysubstr(CONDITION,1,40) $mycountq1 Condition $mycountq2,$sep1, $myquote IN $myquote from DEF_LNKI_P a,DEF_JOB b,DEF_TABLES c $nolock where a.JOB_ID=b.JOB_ID and b.TABLE_ID=c.TABLE_ID and a.TABLE_ID=b.TABLE_ID and a.CONDITION like $cond ";              
        dosql();
        system "type c:\\temp\\emminer.report.01.out >> c:\\temp\\emminer.report.01.out.inout";

gouts: $cond="$glp";
        $cond="$myquote$cond$mypat$myquote";
        print temp "select DATA_CENTER,$sep,SCHED_TABLE,$sep1,JOB_NAME $mycountq1 Jobname $mycountq2,$sep1, $mysubstr(CONDITION,1,40) $mycountq1 Condition $mycountq2,$sep1, $myquote OUT $myquote  from DEF_LNKO_P a,DEF_JOB b,DEF_TABLES c  $nolock where a.JOB_ID=b.JOB_ID and  b.TABLE_ID=c.TABLE_ID and a.TABLE_ID=b.TABLE_ID and a.CONDITION like $cond   ";               
        dosql();
        system "type c:\\temp\\emminer.report.01.out >> c:\\temp\\emminer.report.01.out.inout";

nextgps:
      }						# end of while (<gps>) loop
     
         close gps;

         open (outt,">c:\\temp\\emminer.report.01.out");
         open (intome,"<c:\\temp\\emminer.report.01.out.inout");
         
         $head1=0;				# some code to decide to put headings or just data lines (need to revisit to see if needed)
         $head2=0;
         $nomoreheads=0;
         
         while (<intome>)
          {
          	$i2=index($_,"-----");
          	$i1=index($_,"DATA_CENTER");
          	if ($i1 > -1)
          	  {
          	     if ($head1 == 0) {print outt "$_";$head1=1;goto nextgpsin;}
          	     goto nextgpsin;
          	  }
          	if ($i2 > -1)
          	  {
          	     if (($head1 == 1) && ($nomoreheads == 0)) {print outt "$_";$nomoreheads=1;goto nextgpsin;};
          	     goto nextgpsin;
          	  }
          	print outt "$_";
nextgpsin:          	     
          }
          
          close outt;
          close intome;
         
         putsheet();  				#puts the sheet for jobs with in or out that were global conditions
         
   print "\n    $querycnt/$querytot      Alerts\n";
                  
  print temp "select count(*) $mycountq1#Alerts by Handled status$mycountq2,$sep,HANDLED from ALARM $nolock GROUP BY HANDLED ORDER BY HANDLED   ";
  $querycnt=$querycnt+1;
  dosql();
  $current_sheet="Alerts";
  putsheet();


  print "    $querycnt/$querytot      Shouts\n";
                  
  print temp "select count(*) $mycountq1   #Jobs with Shout      $mycountq2,$sep,WHEN_COND from DEF_SHOUT $nolock GROUP BY WHEN_COND ORDER BY WHEN_COND  ";
  $querycnt=$querycnt+1;
  dosql();
  $current_sheet="Shouts";
  putsheet();
  
  print "    $querycnt/$querytot      Command Line queries\n";
  
  $tsk="Command";
  $tsk="$myquote$tsk$myquote";
#  print temp "select CMD_LINE from DEF_JOB where TASK_TYPE=$tsk $nolock ";
  if ($dbtype eq "S")
    {
      print temp "select CMD_LINE from DEF_JOB where TASK_TYPE=$tsk $nolock ";
    }
  elsif (($dbtype eq "M") || ($dbtype eq "E"))
    {
      print temp "select CMD_LINE from DEF_JOB $nolock where TASK_TYPE=$tsk  ";
    }
  $querycnt=$querycnt+1;
  dosql();
  $current_sheet="CMDLine";
  parcecmds();
  putsheet();
  

  print "    $querycnt/$querytot      On statements\n";
  
  $statement="Statement";
  $statement="$mycountq1$statement$mycountq2";                  
  print temp "select count(*) $mycountq1   #Jobs On Statement    $mycountq2,$sep,$mysubstr(STMT,1,25) $statement from DEF_ON $nolock GROUP BY STMT ORDER BY STMT ";
  $querycnt=$querycnt+1;
  dosql();
  $current_sheet="On Stmt";
  putsheet();

  print "    $querycnt/$querytot      By Weekly Days string\n";
                                      
  print temp "select count(*) $mycountq1   #Weekly days strings$mycountq2,$sep,W_DAY_STR from DEF_JOB $nolock GROUP BY W_DAY_STR  "; 
  $querycnt=$querycnt+1;
  dosql();
  $current_sheet="WDays str";
  putsheet();

  print "    $querycnt/$querytot      By Monthly Days string\n";
                                      
  print temp "select count(*) $mycountq1   #Monthly days strings$mycountq2,$sep,DAY_STR from DEF_JOB $nolock GROUP BY DAY_STR  "; 
  $querycnt=$querycnt+1;
  dosql();
  $current_sheet="Days str";
  putsheet();

  print "    $querycnt/$querytot      By Monthly calendars\n";
                                      
  print temp "select count(*) $mycountq1   #Jobs with Monthly Cal$mycountq2,$sep,DAYS_CAL from DEF_JOB $nolock GROUP BY DAYS_CAL  "; 
  $querycnt=$querycnt+1;
  dosql();
  $current_sheet="With Monthly cal";
  putsheet();

  print "    $querycnt/$querytot      By Weekly calendars\n";
                  
  print temp "select count(*) $mycountq1   #with Weekly Cal      $mycountq2,$sep,WEEKS_CAL from DEF_JOB $nolock GROUP BY WEEKS_CAL ORDER BY WEEKS_CAL  ";
  $querycnt=$querycnt+1;
  dosql();
  $current_sheet="With Weekly cal";
  putsheet();

  print "    $querycnt/$querytot      By Holiday calendars\n";
                                     
  print temp "select count(*) $mycountq1   #with Holiday Cal     $mycountq2,$sep,CONF_CAL from DEF_JOB $nolock GROUP BY CONF_CAL ORDER BY CONF_CAL "; 
  $querycnt=$querycnt+1;
  dosql();
  $current_sheet="With Holiday cal";
  putsheet();

  print "    $querycnt/$querytot      Calendars by Data Center\n";
  print "               --> identify Highest year defined for each calendar and any duplicate calendars\n";
  print "               ";
               
  print temp "select DATA_CENTER $mycountq1 Data Center Name        $mycountq2,$sep,CALENDAR,$sep1,PERIODIC from  DF_CALENDAR $nolock GROUP BY DATA_CENTER,CALENDAR,PERIODIC ORDER BY DATA_CENTER,CALENDAR  ";  
  $querycnt=$querycnt+1;
  dosql();
  $cal_count=0;

  open (resultsin,"<c:\\temp\\emminer.report.01.out");
  while (<resultsin>)
    {
    	$l1=length($_);
    	if ($l1 > 2)
    	   {
    	     $cal_count=$cal_count+1;
    	   }
    }
  close resultsin;
  $cal_count=$cal_count-2;	# subtract 2 lines for the headings that SQL produced
  $cal_upd_interval=$cal_count/10;
  print "--> total of $cal_count calendars to process, will show % of completion on next line\n";
  print "                ";
  $current_sheet="Cal by DC";
  $rescal="yes";  				# to signal code which adds highest year for that calendar to the spreadsheet
  putsheet();
  $rescal="no";					# turn off that signal
 
#--------------------------------------------------------------------------------------------------------------
# now traverse the array holding all dup calendar information and put onto existing worksheet with calendar info
#--------------------------------------------------------------------------------------------------------------

  $x=0;
  $sheet->Cells(1,8)->{Value}="duplicate calendars";
  $sheet->Cells(1,4)->{Value}="Highest Yr defined";
  $sheet->Cells(1,5)->{Value}="desc";
  $sheet->Cells(1,6)->{Value}="days 1-183";
  $sheet->Cells(1,7)->{Value}="days 184-366";
  foreach $vx (@calname)
    {
    	$x=$x+1;
    	$row=$x+1;
    	$sheet->Cells($row,8)->{Value}="@caldupno[$x] @caldup[$x]";
    }

  print "\n    $querycnt/$querytot      Application type\n";
                     
  print temp "select count(*) $mycountq1#Jobs by Application type$mycountq2,$sep,APPL_TYPE from DEF_JOB $nolock GROUP BY APPL_TYPE ORDER BY APPL_TYPE  ";
  $querycnt=$querycnt+1;
  dosql();
  $current_sheet="App Type";
  putsheet();
  

#  print "\n    $querycnt/$querytot      ODATs \n";
                     
#  print temp "select count(*) $mycountq1# of conditions$mycountq2,$sep,ODAT from DEF_JOB $nolock GROUP BY ODAT ORDER BY ODAT  ";
#  $querycnt=$querycnt+1;
#  dosql();
#  $current_sheet="App Type";
#  putsheet();
  
    print "    $querycnt/$querytot      Counting tables by Data Center\n";
                   
  print temp "select count(*) $mycountq1   #Sched tables per DC  $mycountq2,$sep,DATA_CENTER from DEF_TABLES $nolock GROUP BY DATA_CENTER   $go";
  $querycnt=$querycnt+1;
  dosql();
  $current_sheet="Tbls per DC";
  putsheet();

  print "    $querycnt/$querytot      Jobs per table\n";
      
  print temp "select count(*) $mycountq1   #Jobs per table by DC $mycountq2,$sep,DEF_TABLES.SCHED_TABLE,$sep1,DEF_TABLES.DATA_CENTER ";
  print temp "from DEF_JOB,DEF_TABLES $nolock where DEF_JOB.TABLE_ID=DEF_TABLES.TABLE_ID GROUP BY DEF_TABLES.DATA_CENTER,";
  print temp "DEF_TABLES.SCHED_TABLE ORDER BY DEF_TABLES.DATA_CENTER,DEF_TABLES.SCHED_TABLE   $go"; 
  $querycnt=$querycnt+1;
  dosql();
  $current_sheet="tbl-dc";
  putsheet();
    
  print "    $querycnt/$querytot      Task type\n";
  
  print temp "select count(*) $mycountq1    Jobs  per Task Type   $mycountq2,$sep,TASK_TYPE from DEF_JOB $nolock GROUP BY TASK_TYPE  $go";
  dosql(); 
  $current_sheet="Tasktype";
  putsheet();
  $querycnt=$querycnt+1;
   
  print "    $querycnt/$querytot      Group\n";
 
  print temp "select count(*) $mycountq1   #Jobs per Group       $mycountq2,$sep,GROUP_NAME from DEF_JOB $nolock GROUP BY GROUP_NAME ORDER BY GROUP_NAME $go";
  $querycnt=$querycnt+1;
  dosql();
  $current_sheet="Grp";
  putsheet();
  
  print "    $querycnt/$querytot      Application\n";

  print temp "select count(*) $mycountq1   #Jobs per Application $mycountq2,$sep,APPLICATION from DEF_JOB $nolock GROUP BY APPLICATION ORDER BY APPLICATION   $go";
  $querycnt=$querycnt+1;
  dosql();
  $current_sheet="App";
  putsheet();

  print "    $querycnt/$querytot      EM users\n";
  
  if ($emver ne "6.1.3")
   {  
   print temp "select $mysubstr(USERNAME,1,20) $mycountq1      EM Users           $mycountq2,$sep,$mysubstr(USERFULLNAME,1,30) $mycountq1 Name$mycountq2,$sep,ISGROUP $mycountq1 GROUP 1=yes$mycountq2,$sep,PASSEXPIREDAYS $mycountq1 Pswd Exp days $mycountq2,$sep,PASSEXPIREDATE $mycountq1 Pswd Exp dt $mycountq2,$sep,ISPASSEXPIRENEXTLOGON $mycountq1 Exp Next Log $mycountq2,$sep,ISACCOUNTLOCKED $mycountq1 Locked $mycountq2,$sep,ACCOUNTLOCKDATE $mycountq1 Locked Dt $mycountq2,$sep,ACCOUNTLOCKORIGINATOR $mycountq1 Lockedby $mycountq2 from GENERALAUTHORIZATIONS $nolock ORDER BY USERNAME   $go \n";                        
   }
  else
   {
   print temp "select $mysubstr(USERNAME,1,20) $mycountq1      EM Users           $mycountq2,$sep,$mysubstr(USERFULLNAME,1,30) $mycountq1 Name$mycountq2,$sep1,ISGROUP $mycountq1 GROUP 1=yes$mycountq2 from GENERALAUTHORIZATIONS $nolock ORDER BY USERNAME   $go \n";                        
   }
   
  $querycnt=$querycnt+1;
  dosql();
  
# now for each user, also show the groups they belong to

  system "copy c:\\temp\\emminer.report.01.out c:\\temp\\emminer.report.emusers \> c:\\temp\\emminer.temp.misclog";
  open (userout,">c:\\temp\\emminer.report.emusers2");
  close userout;					# this was opened/closed to empty the file
  open (userin,"<c:\\temp\\emminer.report.emusers");
  open (userout,">>c:\\temp\\emminer.report.emusers2");
  while (<userin>)
   {
     $i1=index($_,"GROUP 1=yes");
     $i2=index($_,"---------------");
     $i3=length($_);
     $i4=index($_,"affected");
 
     chop;
     print userout ("$_ :-:");
     @userinrec = split(/:-:/,"$_");
     $username=@userinrec[0];
     $username =~ s/^\s+//;
     $username =~ s/\s+$//;        #remove leading & trailing blanks
     
     open (temp,">c:\\temp\\emminer.report.01");
     print temp "select USERGROUP from USERSGROUPS where USERNAME=$myquote$username$myquote";
     dosql();
     open (groupinfo,"<c:\\temp\\emminer.report.01.out");
     while (<groupinfo>)
     {
     	chop;
     	$i1=index($_,"USERGROUP");
        $i2=index($_,"---------------");
        $i3=index($_,"affected");
        $i4=length($_);
        if (($i1 > -1) || ($i2 > -1) || ($i3 > -1) || ($i4 < 2)) {goto nextgr;}
        $_ =~ s/^\s+//;
        $_ =~ s/\s+$//;        #remove leading & trailing blanks        
        print userout " $_, ";
nextgr:        
     }
     print userout "\n";
     close groupinfo;
     
nextinrec:
   } #end of while userinfo
   close userin;

  system "erase c:\\temp\\emminer.report.01.out";
  close userout;
  system "rename c:\\temp\\emminer.report.emusers2 emminer.report.01.out";  
  $current_sheet="EM Users";
  putsheet();
 
   
  print "    $querycnt/$querytot      Agent (shows all agents defined in job definitions)\n";
 
  print temp "select count(*) $mycountq1   #Jobs per agent       $mycountq2,$sep,NODE_ID from DEF_JOB $nolock GROUP BY NODE_ID ORDER BY NODE_ID $go";
  $querycnt=$querycnt+1;
  dosql();
  $current_sheet="Agent";
  $agtprogress=0;					# used as counter to indicate visually "dots" to the screen to show activity
  $resolveip="yes";				# signal flag to activate code to resolve ip addresses
  print "               --> resolving all agent host and showing its IP address also\n";
  $agt_count=0;

  open (resultsin,"<c:\\temp\\emminer.report.01.out");
  while (<resultsin>)
    {
    	$l1=length($_);
    	$i1=index($_,"NULL");
    	if  (($l1 > 2) && ($i1 == -1))
    	   {
    	     $agt_count=$agt_count+1;
    	   }
    }
  close resultsin;
  $agt_count=$agt_count-3;	# subtract 2 lines for the headings that SQL produced and 1 to represent no agent given
  $agt_upd_interval=$agt_count/10;
#print "debug agt_upd_interval=$agt_upd_interval which was $agt_count / 10\n";
  print "               --> total of $agt_count agents to process, will show % of completion on next line\n";
  print "                ";  
  putsheet();
  $resolveip="no";  				# turn off signal flag
 
  print "\n    $querycnt/$querytot      Historical Job Counts for EM (will take a moment)\n";

#------- first find all the old network AJF tables still in the EM DB and sort so "oldest daily" comes first

  $v1="A%JOB";
  $v1="$myquote$v1$myquote";
  if ($dbtype eq "O")
     {
       print temp "select TABLE_NAME from CAT $nolock where TABLE_NAME like $v1 and TABLE_TYPE='TABLE'  ";
     }       
  else
     {      
       print temp "select $mysubstr(name,1,20) $mycountq1 name $mycountq2 from sysobjects $nolock where name like $v1 order by name desc  ";
     }

  $querycnt=$querycnt+1;
  dosql();

  system "copy c:\\temp\\emminer.report.01.out c:\\temp\\emminer.report.02 >c:\\temp\\emminer.report.01.out.copy.txt";

  open (ajfin,"<c:\\temp\\emminer.report.02");

  $olddc="";					# give this variable an original value of nothing
  
#------- then while processing, each "new date" found indicates most recent download.  process it and skip to next "day change"

  while (<ajfin>)
    {
 
    chop;
    $_ =~ s/^\s+//;$string =~ s/\s+$//;        #remove leading & trailing blanks
    $ltest=length($_);
    $ii1=index($_,"TABLE_NAME");
    $ii2=index($_,"ffected");
    $ii3=index($_,"elected");
    $test=substr($_,1,5);
 
    if (($test eq "ame  ") || ($test eq "-----") || ($ltest < 2) || ($ii1 > -1) || ($ii2 > -1) || ($ii3 > -1)) {goto endafjin;}
 
    $dbtbl=substr($_,0,15);
    $dt=substr($_,1,6);
    $mo=substr($_,3,2);
    $yr=substr($_,1,2);
    $day=substr($_,5,2);
    $dc=substr($_,7,3);

    if (($dc ne $olddc) || ($dt ne $olddt))
       {
         $dc1=$dc;
         $dc1="$myquote$dc$myquote";
       	 print temp "select count(*) $mycountq1      # jobs             $mycountq2,$sep,$myquote $mo-$day-20$yr $myquote $mycountq1     date                $mycountq2,$sep1,DATA_CENTER from $dbtbl,COMM $nolock WHERE COMM.CODE=$dc1 GROUP BY DATA_CENTER \n";
       	 print temp "$go";
       	 print temp "$myblank";
       	 print temp "$go";
       	 $olddc=$dc;
       	 $olddt=$dt;
       }
endafjin:    	
    }
  close ajfin;
 
  dosql();
 
  system "copy c:\\temp\\emminer.report.01.out c:\\temp\\emminer.report.02 >c:\\temp\\emminer.report.01.out.copy.txt";
  open dccountin,("<c:\\temp\\emminer.report.02");
  open dccountout,(">c:\\temp\\emminer.report.01.out");

  $head1="0";					# again some code originally to decide to add headers or not...
  $head2="0";
  $sum=0;
  $olddt="";
  $firstsum=0;
 
  while (<dccountin>)
   {
     @colarray2 = split(/:-:/,"$_");

     chop;
     $l1=length("$_");
     $i1=index($_,"DATA_CENTER");
     $i2=index($_,"-------");
     $i3=index($_," affected");
     $i4=index($_," selected");
     $i5=index($_,"TABLE_NAME");
     $i6=index($_,"no rows");
  
     if (($i1 > -1) && ($head1 eq "0")) 
        {
           $head1="1";
           print dccountout "$_\n";
        }
     if (($i1 > -1) && ($head1 eq "1")) {goto endcountin;}
     if (($i2 > -1) && ($head2 eq "1")) {goto endcountin;}
     if (($i2 > -1) && ($head2 eq "0")) 
        {
           $head2="1";
           print dccountout "$_\n";
        }
     if (($i2 > -1) || ($i4 > -1) || ($i3 > -1) || ($l1 < 2) || ($i5 > -1) || ($i6 > -1)) {goto endcountin;}

     $dt=@colarray2[1];
     $ct=@colarray2[0];
  
     if ($olddt ne $dt)
        {
	    if ($sum > $maxltask) {$overunder="over Max Tasks"}; 	
        	
            if ($firstsum < 1) 
               {
               	  $olddt=$dt;
               	  $firstsum=1;
               	  $biggest1day=$ct;
               	  goto ptb;
               }
               
            $overunder=" ";
          
            print dccountout ":-::-::-:$olddt--total  $sum    $overunder\n";
            
            if ($biggest1day < $sum)
             {
             	$biggestdate=$dt;
             	$biggest1day = $sum;
             }
             
            $olddt=$dt;
            $sum=0;
        }
ptb:        
     $sum=$sum+$ct;
     print dccountout "$_\n";
 
endcountin:
   }

            if ($biggest1day < $sum) 
              {
              	$biggest1day = $sum;
              	$biggestdate=$dt;
              }

  if ($sum > $maxltask) {$overunder="over Max Tasks"}; 	

  print dccountout ":-::-::-:$dt  --total $sum    $overunder\n";
  print dccountout ":-::-::-:\n";
  print dccountout ":-::-:HIGHWATER occurred:-:$biggestdate$biggest1day\n"; 
  print dccountout ":-::-:Max Licensed Tasks:-::-::-::-:$maxltask\n"; 
  close dccountin;
  close dccountout;
  
  $current_sheet="Jobs in EM AJF";
  putsheet();
 
  print "\n    $querycnt/$querytot      Total jobs\n";

  print temp "select count\(*\) $mycountq1 Total jobs def in $server $mycountq2 from DEF_JOB $nolock $go";
  dosql();
  $current_sheet="Misc";
  $miscrow=$miscrow+1;					# keep track of what row we are on in the "misc" worksheet
  $miscsheet=1;
  $misccol=1;
  putsheet();
  $miscsheet=0;
 
  $querycnt=$querycnt+1;
  
  print "    $querycnt/$querytot      Cyclic\n";

  $cyc="1";
  $cyc="$myquote$cyc$myquote";
  print temp "select count(*) $mycountq1  # Cyclic Jobs      $mycountq2 from DEF_JOB $nolock where CYCLIC=$cyc  $go";
  $querycnt=$querycnt+1;
  dosql();
  $current_sheet="Misc";
  $miscsheet=1;
  $miscrow=$miscrow+1;
  $misccol=1;  
  putsheet();
  $miscsheet=0;
#  $querycnt=$querycnt+1;
  
  print "    $querycnt/$querytot      Jobs with Autoedits\n";

  print temp "select count(*) $mycountq1  # Jobs w sysouts      $mycountq2,JOB_ID,TABLE_ID from DEF_SETVAR $nolock GROUP BY TABLE_ID,JOB_ID ORDER BY TABLE_ID,JOB_ID $go";
  $querycnt=$querycnt+1;
  dosql();
  $current_sheet="Misc";
  $miscsheet=1;
  $miscrow=$miscrow+1;
  $misccol=1; 
  
  # now must adjust file so a single value will be placed onto the misc sheet (not each job with its count of autoedits)
  close temp;						# was reopened for output in dosql
  open (aein,"<c:\\temp\\emminer.report.01.out");

  while (<aein>)
   {
     $i1=index($_,"TABLE_ID");
     $i2=index($_,"------");
     $i3=length($_);
     $i4=index($_,"affected");
     if (($i1 < 0) && ($i2 < 0) && ($i4 < 0) && ($i3 > 10)) {$aejobs++;}
   } 
  close aein;
      
  open (temp,">c:\\temp\\emminer.report.01.out");
  print temp "  # Jobs with Autoedits \n";
  print temp "-----------------\n";
  print temp "    $aejobs\n";
  close temp;
  putsheet();
  
  $miscsheet=0;
  open (temp,">c:\\temp\\emminer.report.01");		# normally this done in dosql but because of above section needs redone
  print temp "$mysqlpre1";
  print temp "$mysqlpre2";   
  
  
  print "    $querycnt/$querytot      Critical\n";

  $crit="1";
  $crit="$myquote$crit$myquote";
  print temp "select count(*) $mycountq1  # Critical Jobs $mycountq2 from DEF_JOB $nolock where CRITICAL=$crit   $go";
  dosql();
  $current_sheet="Misc";
  $miscsheet=1;
  $miscrow=$miscrow+1;
  $misccol=1;  
  putsheet();
  $miscsheet=0;
  $querycnt=$querycnt+1;

  print "    $querycnt/$querytot      Confirm required\n";

  $conf="1";
  $conf="$myquote$conf$myquote";
  print temp "select count(*) $mycountq1  # with Confirm required $mycountq2 from DEF_JOB $nolock where CONFIRM_FLAG=$conf   $go";
  dosql();
  $current_sheet="Misc";
  $miscsheet=1;
  $miscrow=$miscrow+1;
  $misccol=1;  
  putsheet();
  $miscsheet=0;
  $querycnt=$querycnt+1;
  print "    $querycnt/$querytot      Multi Agent\n";

  $ma="Y";
  $ma="$myquote$ma$myquote";
  print temp "select count(*) $mycountq1  # Multi-Agent Jobs $mycountq2 from DEF_JOB $nolock where MULTY_AGENT=$ma   $go";
  dosql();
  $current_sheet="Misc";
  $miscsheet=1;
  $miscrow=$miscrow+1;
  $misccol=1;  
  putsheet();
  $miscsheet=0;
  $querycnt=$querycnt+1;
  
  print "    $querycnt/$querytot      Active From & Until\n";
 
  $ttoday="Y";
  $ttoday="$myquote$today$myquote";
  print temp "select count(*) $mycountq1  # Active in Future $mycountq2 from DEF_JOB $nolock where ACTIVE_FROM > $ttoday  $go";
  dosql();
  $current_sheet="Misc";
  $miscsheet=1;
  $miscrow=$miscrow+1;
  $misccol=1;  
  putsheet();
  $nada="$myquote$myquote";
  print temp "select count(*) $mycountq1  # Active Until has Past $mycountq2 from DEF_JOB $nolock where ACTIVE_TILL < $ttoday and ACTIVE_TILL > $nada  $go";
  dosql();
  $current_sheet="Misc";
  $miscsheet=1;
  $miscrow=$miscrow+1;
  $misccol=1;  
  putsheet();
  
  
  $miscsheet=0;
  
  
  $querycnt=$querycnt+1;  
  

  print "    $querycnt/$querytot      Pre or Post commands\n";
                  
  print temp "select count(*) $mycountq1  # Using PRE or POST CMD$mycountq2 from DEF_SETVAR $nolock where NAME=$myquote%%PRECMD$myquote or NAME=$myquote%%POSTCMD$myquote ";
  $querycnt=$querycnt+1;
  dosql();
  $current_sheet="Misc";
  $miscsheet=1;
  $miscrow=$miscrow+1;
  $misccol=1;
  putsheet();
  $miscsheet=0;

  print "    $querycnt/$querytot      Retro\n";

  $ret="1";
  $ret="$myquote$ret$myquote";
  print temp "select count(*) $mycountq1  # with Retro      $mycountq2 from DEF_JOB $nolock where RETRO=$ret   $go";
  $querycnt=$querycnt+1;
  dosql();
  $current_sheet="Misc";
  $miscsheet=1;
  $miscrow=$miscrow+1;
  $misccol=1;
  putsheet();
  $miscsheet=0;
  print "  \n";
  close temp;
                                                   
  $sheet = $book->Worksheets("Misc");  
  $sheet->Activate();
  
enddbquery:
}

#======================     parcecmds function   ===========================

sub parcecmds()
{
  if ($debug) {print "--- debug parcecmds routine\n";}
  if (-f "c:\\temp\\emminer.cmdstrings")
     {
     	$donothing=1;
    	# already have a file containing cmd line strings to be searched for
     }
  else
      {
      	print "\nnote: Data Miner is building a default set of strings to scan for within\n";
      	print "      job definitions COMMAND LINE.  You can adjust or add to the strings\n";
      	print "      scanned for by editing the file c:\\temp\\emminer.cmdstrings\n\n";
      	
      	initial_cmdstrings();  #else give the user a default file
      }
      
  system "copy c:\\temp\\emminer.report.01.out c:\\temp\\emminer.report.cmds > c:\\temp\\emminer.report.01.out.copy.txt";
  open (cmds,"<c:\\temp\\emminer.report.cmds");
  open (cmdstrings,"<c:\\temp\\emminer.cmdstrings");
  open (rpt,">c:\\temp\\emminer.report.01.out");	# will be printed later via putsheet


# load up my array which holds the strings to be looked for
  $stringno=0;
  while (<cmdstrings>)
     {
     	chop;
     	$x=substr($_,0,1);
     	if ($x eq "#") {goto nextcmdstring;}
     	@cmdarray_val[$stringno]=$_;
     	$stringno=$stringno+1;
     	
nextcmdstring:
     }
  close cmdstrings;
   	
  while (<cmds>)
   {
     $heading1=index($_,"CMD_LINE");
     $heading2=index($_,"----------------");
     if (($heading1 > -1) || ($heading2 > -1)) {goto nextonea;}
     $stringno=0;
     foreach $cmd (@cmdarray_val)	 
       {
       	 if (index($_,$cmd) > -1)
       	    {
       	    	@cmdarray_cnt[$stringno]=@cmdarray_cnt[$stringno]+1;
       	    }
       	 $stringno=$stringno+1;
       	    
       } #end of foreach cmd array
nextonea:
   }    
# now have finished counting each hit of the string across all command lines, so report them

   print rpt " Count :-: String\n";  
   $stringno=0;
   foreach $cmd (@cmdarray_val)	 
       {
       	 $stringno=$stringno+1;
       	 print rpt "@cmdarray_cnt[$stringno] :-: @cmdarray_val[$stringno] \n";
       } 
   close rpt;
   close cmds;       
}  #end of parcecmds function

#======================     dosql function   ===========================

sub dosql ()
{
 if ($debug) {print "--- debug dosql routine\n";}
	
  print temp "$go$myblank$go$myblank$go"; 
	
  if (($dbtype eq "M") || ($dbtype eq "S") || ($dbtype eq "E"))
     {     	
       print temp "exit\ngo\n";
       close temp;
       if ($debug) {print "debug --- executing $sqlcmd -i c:\\temp\\emminer.report.01 -o c:\\temp\\emminer.report.01.out";}
       system "$sqlcmd -i c:\\temp\\emminer.report.01 -o c:\\temp\\emminer.report.01.out";
     }
  else
     {
       print temp "exit;";
       close temp;
       system " $sqlcmd \@c:\\temp\\emminer.report.01 > c:\\temp\\emminer.report.01.out";
     }
     
#  close scout;
  if ($debug)
     {
     	print "\n\n--debug results of executing sql follow\n";
        system "type c:\\temp\\emminer.report.01\n";
        system "type c:\\temp\\emminer.report.01.out\n";
     }
#  system "type c:\\temp\\emminer.report.01.out >>c:\\temp\\emminer.report";
#  open (scout,">>c:\\temp\\emminer.report");
  open (temp,">c:\\temp\\emminer.report.01");
  print temp "$mysqlpre1";
  print temp "$mysqlpre2";   

}

#======================     putsheet function   ===========================
       
sub putsheet
   {
    if ($debug) {print "--- debug putsheet routine\n";}
   	
    if (($Misc_sheet_created) && ($current_sheet eq "Misc"))
       {
       	 $sheet = $book->Worksheets("$current_sheet");		# just point to the tab sheet if it already exist like Misc will later
         $sheet->Activate();
       	 goto gotmisc; 
       }
	
     if ($current_sheet eq "Misc") {$Misc_sheet_created=1;}
     $sheet = $book->Worksheets->add(); 
     $sheet->Activate();
     $sheet->{Name}="$current_sheet";   
 
gotmisc:    

     open (resultsin,"<c:\\temp\\emminer.report.01.out");
        $row=0;
        $heading="";

        while (<resultsin>)
           {
             chop;
	     $OK_CHARS='-a-zA-Z0-9_.@ *';	# A restrictive list, which
		        			# should be modified to match
					        # as appropriate              

             $l1=length($_);

             if ($current_sheet eq "Globals")
                {
     	        $_ =~ s/^\s+//;
     	        $_ =~ s/\s+$//;        #remove leading & trailing blanks
	        $_ =~s/     //g;
	        $_ =~s/    //g;
	        $_ =~s/   / /g;
	        $_ =~s/  / /g;
	        $_ =~s/ //g;   
                }

             if ($l1 < 2) {goto skipra;}
             $rowsaff1=index($_," affected");
             $rowsaff2=index($_," selected");
             $countline=index($_,"count(*)");
             $substrline=index($_,"substr");
             $dashes=index($_,"------");
             if (($rowsaff1 > -1) || ($rowsaff2 > -1) || ($dashes > -1) || ($countline > -1) || ($substrline > -1)) {goto skipra;};
             
             $row=$row+1;  
           
   	     @colarray = split(/:-:/,"$_");
             $col=0;
        
eachln:	     foreach $xx (@colarray)	# put each member of the array (columns) into individual cells for this row
	       {
	        $col=$col+1;
	        $collen=length($xx);

	        if ($miscsheet)
	           {
                     $sheet->Cells($miscrow,$misccol)->{Value}=$xx;   
 	             $misccol=$misccol+1;
                     $sheet->Columns($col)->{ColumnWidth}=$collen;   
 	           }
	        else
	           {
                     $xxlen=length($xx);
                     if ($xxlen > 200) {$xx=substr($xx,1,50);$collen=55;}
                     $sheet->Cells($row,$col)->{Value}=$xx;   
                     $sheet->Columns($col)->{ColumnWidth}=$collen;         
                   }
               }
                   
resip:      if ($resolveip eq "yes")		# if this is the agent query info, resolve the ip for each agent
               {

                $tname=@colarray[1];
     	        $tname =~ s/^\s+//;
     	        $string =~ s/\s+$//;        #remove leading & trailing blanks
	        $tname  =~s/     //g;
	        $tname  =~s/    //g;
	        $tname  =~s/   / /g;
	        $tname  =~s/  / /g;
	        $tname =~s/ //g;   
	        $riplen1=length($tname);
	        $ripi1=index($tname,"NODE_ID");
	        $ripi2=index($tname,"------");
	        $ripi3=index($tname,"NULL");
	        if (($ripi1 > -1) || ($ripi2 > -1) || ($ripi3 > -1)) {;goto endresip;};
	        if ($riplen1 < 1) {goto endresip;} 
	        
              	$agtprogress=$agtprogress+1;
#     		if ($agtprogress > $agt_upd_interval)
     		if (($agtprogress == $agt_upd_interval) || ($agtprogress > $agt_upd_interval)) 
     		 
     		   {
#print "debug agtprogress of $agtprogress => agt_upd_interval of $agt_upd_interval\n";
     		     $tot_agt_done=$tot_agt_done+$agtprogress;
     		     if ($agt_count > 0)
     		        {
     		         $tot_perc_done=$tot_agt_done/$agt_count*100;
     		        }
     		     else
     		        {
     		         $tot_perc_done=100;
     		        }
                     print "---";
     		     printf ("%3d",$tot_perc_done);
      		     print "$percent";
     		     $agtprogress=0;
     		   }	# show activity by update user on status
     		                		        
	        
                system "ping -n 1 $tname > c:\\temp\\emminer.report.rip.txt";       
                open (ripret,"<c:\\temp\\emminer.report.rip.txt");

                while (<ripret>)
                  {
                     $ripi1=index($_,"Reply from ");
                     $ripi2=index($_,":");
                     if ($ripi1 > -1)
                        {
                           $ip=substr($_,$ripi1+11,$ripi2-$ripi1-11);
                           $sheet->Cells($row,3)->{Value}=$ip;   
                        }
                  }
endresip:
               }  # end of the resip if loop	   
             
             close ripret;

#---------------------------------------------------------------------------------
# if doing the calendar sheet, identify what is the latest year for each calendar             
#---------------------------------------------------------------------------------             
          
rescalendar:    if (($rescal eq "yes") && (@colarray[2] ne "  PERIODIC "))  # skips the heading line from previous query
             {
                open (cal,">c:\\temp\\emminer.report.cal.sql");
                $dcname=@colarray[0];
     	        $dcname =~ s/^\s+//;
     	        $dcname =~ s/\s+$//;        #remove leading & trailing blanks
	        $dcname  =~s/     //g;
	        $dcname  =~s/    //g;
	        $dcname  =~s/   / /g;
	        $dcname  =~s/  / /g;
	        $dcname =~s/ //g;   
                $calname=@colarray[1];
     	        $calname =~ s/^\s+//;
     	        $calname =~ s/\s+$//;        #remove leading & trailing blanks
	        $calname  =~s/     //g;
	        $calname  =~s/    //g;
	        $calname  =~s/   / /g;
	        $calname  =~s/  / /g;
	        $calname =~s/ //g;  	       
	        $callen1=length($dcname);
	        $cali1=index($dcname,"DataCenter");
	        $cali2=index($dcname,"----");
	       
	        if (($cali1 > -1)) {$field4="Highest year defined";goto endrescal;};	# skip header lines
	        if (($cali2 > -1)) {$field4="-------------------";goto endrescal;};	# skip dash lines
	        if ($callen1 < 1) {goto endrescal;} 	# skip blank lines
	       
             	$calprogress=$calprogress+1;
     		if (($calprogress == $cal_upd_interval) || ($calprogress > $cal_upd_interval)) 
     		   {
     		     $tot_cal_done=$tot_cal_done+$calprogress;
     		     if ($cal_count < 1) {$tot_perc_done=100;}		# to be sure no divide by zero
     		     else {$tot_perc_done=$tot_cal_done/$cal_count*100;}
                     print "---";
     		     printf ("%3d",$tot_perc_done);
      		     print "$percent";
     		     $calprogress=0;
     		   }	# show activity by update user on status	       
	       
	        $calname="$myquote$calname$myquote";
	        $dcname="$myquote$dcname$myquote";
                if (($dbtype eq "M") || ($dbtype eq "S") || ($dbtype eq "E"))
                    {
                      print cal "set nocount on  $go";
                    }
                else 
                    {
                      print cal "set pagesize 9999;\n";
                      print cal "set linesize 9000;\n";
                      print cal "set tab off;\n";
                    }
  	        print cal "select YEAR $mycountq1 Calendar Years $mycountq2,$sep,DESCRIPTION,$sep1,DAYS_1 $mycountq1 Days part 1 $mycountq2,$sep1,DAYS_2 $mycountq1 Days part 2 $mycountq2 from DF_YEARS $nolock where CALENDAR=$calname and DATA_CENTER=$dcname order by YEAR DESC   $go";
  	        print cal "$myblank";
  	        print cal "$go";       

  
   	       if (($dbtype eq "M") || ($dbtype eq "S") || ($dbtype eq "E"))
    	       {
     	     	print cal "exit";
       	        close cal;

   	        system "$sqlcmd -i c:\\temp\\emminer.report.cal.sql -o c:\\temp\\emminer.report.cal.sql.out";
       	        if ($debug)
       	           {
       	             print "\n\ndebug in calendar processing sql and results follow\n";
	             system "type c:\\temp\\emminer.report.cal.sql \n";
	             system " type c:\\temp\\emminer.report.cal.sql.out";
	           }
	       }
  	      else
  	       {
  	        print cal "exit;";
      	        close cal;
      	        system " $sqlcmd \@c:\\temp\\emminer.report.cal.sql \> c:\\temp\\emminer.report.cal.sql.out";
    	       }  

               open (rcalret,"<c:\\temp\\emminer.report.cal.sql.out");
               while (<rcalret>)
                  {
                     chop;
                     $cali1=index($_,"Years");
                     $cali2=index($_,"----");
                     $cali3=index($_,"elected");
                     $cali4=index($_,"ffected");		
                     $cali6=index($_,"CALENDAR");
                     $cali5=index($_,"DESCRIPTION");
                     $cali7=index($_,"Days part");
                     $call1=length($_);
                   
                     if ($cali1 > -1) {$field4="Highest Yr Defined";goto nxtyr;}
                     if  (($cali1 > -1) || ($cali7 > -1) || ($cali2 > -1 ) || ($cali5 > -1) || ($call1 < 3) || ($cali3 > -1) || ($cali4 > -1) || ($cali6 > -1)) {goto nxtyr;}
                     if (($cali1 < 0) && ($cali2 < 0) && ($call1 > 3) && ($cali3 < 0) && ($cali4 < 0))
                        {
                         @colarray2 = split(/:-:/,"$_");            
                         $sheet->Cells($row,4)->{Value}=@colarray2[0];
                         if ((@colarray2[0] eq "NULL") || (@colarray2[0] eq "")) {$sheet->Cells($row,4)->{Value}="no yr Defined";};
                         $sheet->Cells($row,5)->{Value}=@colarray2[1];   
                         $sheet->Cells($row,6)->{Value}=@colarray2[2];   
                         $sheet->Cells($row,7)->{Value}=@colarray2[3];   
                         $sheet->Cells($row,8)->{Value}=@colarray2[4];       

dupcheck: 	         $v1=$row-1;
	                 if ($v1 > 0)
	  	           {
		             @calname[$v1]=$calname;
		             $f1="@colarray2[2]@colarray2[3]";
#this cause similar periodic calendars to show as dups that were not            $f1 =~s/ //g;  	       
		             @calval[$v1]="$f1";
		             $checkrow=$v1-1;

		            while ($checkrow > 0)
  			         {
			            if (@calval[$v1] eq @calval[$checkrow])
 			               {
 			               	$prevdups_checkrow=@caldup[$checkrow];
 			               	$prevdups_v1=@caldup[$v1];
 			               	$thisdup=@calname[$checkrow];
      			         	@caldup[$v1]="$prevdups_v1 $thisdup";
      			         	@caldup[$checkrow]="$prevdups_checkrow $calname";
      			         	@caldupno[$checkrow]=@caldupno[$checkrow]+1;
      			         	@caldupno[$v1]=@caldupno[$v1]+1;      	
   			               }	
   		                    $checkrow=$checkrow-1;
  			         }
                            } 
                            goto donethiscal;
nxtyr:
                           }
endrescal:
                      }
             	   
donethiscal:             
                   close rcalret;
                   close cal;                    

             }          

skipra:      
           }
        
doneput:
        close resultsin;
        
    }   		


#======================     getconfig function   ===========================

sub getconfig ()
{
 if ($debug) {print "--- debug getconfig routine\n";}
 if (-e "c:\\temp")		#verify that a temp directory exist or create it
    {
    	#its already there if I take the IF
    }
   else
    {
    	system "mkdir c:\\tempt > c:\\temp\\emminer.report.01.out.copy.txt";
    }

 if (-e "c:\\temp\\emminer.config")
    {
    open (config,"c:\\temp\\emminer.config") || die "Can't open config file c:\temp\emminer.config\n";
    while (<config>)
          {
           chop;
           $fld=substr($_,0,5);
           $val=substr($_,5);
           if ($fld eq "emvr:") { $emver="$val"; }
           if ($fld eq "emal:") { $emal="$val"; }
           if ($fld eq "fpre:") { $fpref="$val";  }
           if ($fld eq "user:") { $emuser="$val";  }
           if ($fld eq "serv:") { $server="$val";  }
	   if ($fld eq "mtsk:") { $maxltask="$val";  }   
	   if ($fld eq "pass:") { $empass="$val";  }           
           if ($fld eq "dbty:") { $dbtype="$val";  }              
          }
     close config;
    }
}

#===========================  getuser_input function =======================
    
sub getuser_input()
   {
getver:
   print "    ---> em version number 6.3, 6.2, or 6.1.3 ($emver):";
   $ans = <STDIN>;                          	#get options
   chop $ans;                               	#remove carrage return
   if (("$ans" eq "" ) && ("$emver" eq "")) { print " --- must enter a version number ---\n"; goto getver; }
   if ("$ans" ne "" ) { $emver=$ans; }

getuser:
   print "    ---> em userid ($emuser):";
   $ans = <STDIN>;                          	#get options
   chop $ans;                               	#remove carrage return
   if (("$ans" eq "" ) && ("$emuser" eq "")) { print " --- must enter a user ---\n"; goto getuser; }
   if ("$ans" ne "" ) { $emuser=$ans; }

getserver:
 
   if (($user_help eq "NO") || ($server ne "")) {goto nohelpmsg;}
   printsyntax();

nohelpmsg:   print "    ---> db server ($server):";
   $ans = <STDIN>;                          	#get options
   chop $ans;                               	#remove carrage return
   if (("$ans" eq "" ) && ("$server" eq "")) { print " --- must enter a user ---\n"; goto getserver; }
   if ("$ans" ne "" ) { $server=$ans; }
 
getdbtype:
   
   print "    ---> db type {m for MSSQL, e for MSDE, s for SYBASE, or o for ORACLE} ($dbtype):";
   $ans = <STDIN>;                          	#get options
   chop $ans;                               	#remove carrage return
   $ans=uc($ans);
   
   if (("$ans" eq "" ) && ("$dbtype" eq "")) { print " --- must enter a user ---\n"; goto getdbtype; }   
   if ("$ans" ne "" ) 
      { $dbtype=$ans; }
 
    if (($dbtype ne "M") && ($dbtype ne "S") && ($dbtype ne "O") && ($dbtype ne "E"))
      {
      	print "\n\nsorry but a dbtype of $dbtype . is not supported.  Try again\n";
      	goto getdbtype;
      }

getMAXtask:
   print "\n";
   print "         If you know are licenced by a max number of task for Control-M \n";
   print "         and you would like this shown on the excel spreadsheet, the routine will \n";
   print "         indicate this number and let you know if your daily usage is over that maximum.  \n";
   print "\n";
   print "         If you are tier licenced or do not know or care for this indicator, just enter\n";
   print "         something like ";
   print '"Not Measured"';
   print " at the prompt\n\n"; 
   print "    ---> max licensed tasks ($maxltask):";
   $ans = <STDIN>;                          	#get options
   chop $ans;                               	#remove carrage return
   if (("$ans" eq "" ) && ("$maxltask" eq "")) { print " --- enter max licensed value ---\n"; goto getMAXtask; }
   if ("$ans" ne "" ) { $maxltask=$ans; }

   

getpswd:
   print "    ---> em user password ($empass):";
   $ans = <STDIN>;                        	#get options
   chop $ans;                              	#remove carrage return
   if ("$ans" ne "" ) { $empass=$ans; }

   system "cls";			  	#clear screen to hide entered password
   print ("emminer.pl starting Data Mining ... \n\n");


#   open (scout , ">c:\\temp\\emminer.report");	#open the temp report text file
}    
    
#===========================  startexcel function =======================
    
sub startexcel ()
{
	if ($debug) {print "--- debug startexcel routine\n";}
#------------------------------------------------------
# use existing instance if Excel is already running
#------------------------------------------------------

        eval {$ex = Win32::OLE->GetActiveObject('Excel.Application')};
        die "Excel not installed" if $@;
        unless (defined $ex) {
            $ex = Win32::OLE->new('Excel.Application', sub {$_[0]->Quit;})
                    or die "Oops, cannot start Excel";
        }
           
  $book = $ex->Workbooks->Add; 			# get a new excel workbook, will name when saving
}

#===========================  cleanup function =======================
    
sub cleanup ()
{
#------------------------------------------------
# now save the excel spreadsheet and close excel
#------------------------------------------------

     $book->SaveAs( "c:\\temp\\$excelfile" ); 	# save/close the excel spreadsheet
     $ex->ActiveWorkbook->Close(0);		# close current workbook
     $ex->Quit(); 				# quit excel        
     system ("erase c:\\temp\\emminer.temp*");
     system ("erase c:\\temp\\emminer.report*"); 
     if ($silent) {goto Fin;}			# don't invoke Excel if were running silently
Doexcel:  system ("start c:\\temp\\$excelfile");
Fin: exit 0;
}

#------------------------------------------------------------------------------------------
#update the config file with current settings (saves ftp info for next time and naming info)
#------------------------------------------------------------------------------------------

sub updconfig ()
{
      open (config,"> c:\\temp\\emminer.config") || die "Can't open config file c:\temp\emminer.config\n";
      print config "emal:$pemal\n";
      print config "fpre:$fpref\n";
      print config "user:$emuser\n";
      print config "serv:$server\n";
      print config "dbty:$dbtype\n";
      print config "emvr:$emver\n";
      print config "mtsk:$maxltask\n";
      print config "pass:$empass\n";
      close config;

}
    
#-----------------------------------
# general routine to get system time
#-----------------------------------

sub gettime()
    {
     ($sec,$min,$hour,$mday,$mon,$year,$wday,$yday,$isdst)=localtime(time);
     if (length($mday) < 2) {$tsec="0$mday";}      
     if (length($min) < 2) {$min="0$min";}  
     if (length($sec) < 2) {$sec="0$sec";}  
     if (length($hour) < 2) {$hour="0$hour";} 
     $year=$year+1900;     
     $mon=$mon+1;
     if ($mon < 10) {$mon="0$mon";}             
    }

#-----------------------------------
# routine to set any specific column widths to values other than the default which is the entire column width from the DB
#-----------------------------------
    
sub override_colwidth ()
    {
       if ($debug) {print "--- debug override_colwidth routine\n";}
       $sheet = $book->Worksheets("SNMP");
       $sheet->Columns(1)->{ColumnWidth}=30;    	
       $sheet->Columns(2)->{ColumnWidth}=60;        
       
       $sheet = $book->Worksheets("Components");
       $sheet->Columns(1)->{ColumnWidth}=40;    	
       $sheet->Columns(2)->{ColumnWidth}=33;
       $sheet->Columns(3)->{ColumnWidth}=41;
       $sheet->Columns(5)->{ColumnWidth}=159;        
             
       $sheet = $book->Worksheets("Globals");
       $sheet->Columns(1)->{ColumnWidth}=18;
              
       $sheet = $book->Worksheets("Tbls per User Daily");
       $sheet->Columns(1)->{ColumnWidth}=24;
              
       $sheet = $book->Worksheets("Jobs in EM AJF");
       $sheet->Columns(1)->{ColumnWidth}=8;
       $sheet->Columns(2)->{ColumnWidth}=12;
       $sheet->Columns(3)->{ColumnWidth}=20;
       $sheet->Columns(4)->{ColumnWidth}=10;
       $sheet->Columns(5)->{ColumnWidth}=10;
       $sheet->Columns(6)->{ColumnWidth}=10;
       $sheet->Columns(7)->{ColumnWidth}=10;

       $sheet = $book->Worksheets("EM Users");
       $sheet->Columns(1)->{ColumnWidth}=25;
       $sheet->Columns(2)->{ColumnWidth}=65;
       $sheet->Columns(3)->{ColumnWidth}=12;
       $sheet->Columns(4)->{ColumnWidth}=25;
       $sheet->Columns(5)->{ColumnWidth}=20;
       $sheet->Columns(6)->{ColumnWidth}=15;
       $sheet->Columns(7)->{ColumnWidth}=7;
       $sheet->Columns(8)->{ColumnWidth}=15;
                
               

       
 
       $sheet = $book->Worksheets("Cal by DC");
       $sheet->Columns(3)->{ColumnWidth}=10;
       $sheet->Columns(6)->{ColumnWidth}=11;
       $sheet->Columns(7)->{ColumnWidth}=11;
       $sheet->Columns(8)->{ColumnWidth}=40;       
                           
       $sheet = $book->Worksheets("Agent");
       $sheet->Columns(1)->{ColumnWidth}=20;       
       $sheet->Columns(2)->{ColumnWidth}=50;       
       $sheet->Columns(3)->{ColumnWidth}=20; 
       
       $sheet = $book->Worksheets("Misc");
       $sheet->Columns(1)->{ColumnWidth}=50;       

       $sheet = $book->Worksheets("CMDLine");
       $sheet->Columns(1)->{ColumnWidth}=10; 
       $sheet->Columns(2)->{ColumnWidth}=20; 
#       $sheet->sort_data('CMDLine',0,'DESC');	# hmmm that didnt work, may revisit this
                                  
    }
    
    
#-----------------------------------    
#   initial_cmdstrings function 
#-----------------------------------
  
sub initial_cmdstrings ()
{
    if ($debug) {print "--- debug initial_cmdstrings routine\n";}
    open (cmdstrings,">c:\\temp\\emminer.cmdstrings");
# agent utilities
    print cmdstrings "# any line starting with a # will be ignored in the report so you can\n";
    print cmdstrings "# turn off the reporting of any particular string by putting a # in col 1\n";
    print cmdstrings "#----------------------------------------------------------\n";
    print cmdstrings "_exit\n";
    print cmdstrings "_sleep\n";
    print cmdstrings "ag_ping\n";
    print cmdstrings "#ag_diag_comm\n";
    print cmdstrings "ctmag\n";
    print cmdstrings "ctmcontb\n";
    print cmdstrings "ctmcreate\n";
    print cmdstrings "ctmfw\n";
    print cmdstrings "ctmpwd\n";
    print cmdstrings "ctmwincfg\n";
    # em utilities
    print cmdstrings "cli\n";
    print cmdstrings "copydefcal\n";
    print cmdstrings "copydefjob\n";
    print cmdstrings "defcal\n";
    print cmdstrings "defjob\n";
    print cmdstrings "defjobconvert\n";
    print cmdstrings "deftable\n";
    print cmdstrings "deldefjob\n";
    print cmdstrings "duplicatedefjob\n";
    print cmdstrings "exportdefcal\n";
    print cmdstrings "exportdefjob\n";
    print cmdstrings "exportdeftable\n";
    print cmdstrings "updatedef\n";
    print cmdstrings "util\n";
    print cmdstrings "check_gtw\n";
    print cmdstrings "ctl\n";
    # control-m server utilities
    print cmdstrings "ctm_agstat\n";
    print cmdstrings "ctm_backup_bcp\n";
    print cmdstrings "ctm_menu\n";
    print cmdstrings "ctm_restore_bcp\n";
    print cmdstrings "ctmagcln\n";
    print cmdstrings "ctmcalc_date\n";
    print cmdstrings "#ctmcontb\n";
    print cmdstrings "ctmcpt\n";
    print cmdstrings "#ctmcreate\n";
    print cmdstrings "ctmdbbcl\n";
    print cmdstrings "ctmdbcheck\n";
    print cmdstrings "ctmdbopt\n";
    print cmdstrings "ctmdbrst\n";
    print cmdstrings "ctmdbspace\n";
    print cmdstrings "ctmdbtrans\n";
    print cmdstrings "ctmdbused\n";
    print cmdstrings "ctmdefine\n";
    print cmdstrings "ctmdiskspace\n";
    print cmdstrings "ctmcheckmirror\n";
    print cmdstrings "ctmexdef\n";
    print cmdstrings "ctmgetcm\n";
    print cmdstrings "ctmgrpdef\n";
    print cmdstrings "ctmjsa\n";
    print cmdstrings "ctmkilljob\n";
    print cmdstrings "ctmldnrs\n";
    print cmdstrings "ctmloadset\n";
    print cmdstrings "ctmlog\n";
    print cmdstrings "ctmnodegrp\n";
    print cmdstrings "ctmordck\n";
    print cmdstrings "ctmorder\n";
    print cmdstrings "ctmpasswd\n";
    print cmdstrings "ctmping\n";
    print cmdstrings "ctmpsm\n";
    print cmdstrings "ctmrpln\n";
    print cmdstrings "ctmruninf\n";
    print cmdstrings "ctmshout\n";
    print cmdstrings "ctmshtb\n";
    print cmdstrings "ctmspdiag\n";
    print cmdstrings "ctmstats\n";
    print cmdstrings "ctmstvar\n";
    print cmdstrings "ctmsuspend\n";
    print cmdstrings "ctmsys\n";
    print cmdstrings "ctmudchk\n";
    print cmdstrings "ctmudlst\n";
    print cmdstrings "ctmudly\n";
    print cmdstrings "ctmvar\n";
    print cmdstrings "ctmwhy\n";
    print cmdstrings "dbversion\n";
    print cmdstrings "ecactltb\n";
    print cmdstrings "ecaqrtab\n";
    close cmdstrings;   
} #end of initial_cmdstrings function

#-----------------------------------
#    sub testdb 
#-----------------------------------

sub testdb ()
{
  if ($debug) {print "--- debug testdb routine\n";}
	
  if (($dbtype eq "M") || ($dbtype eq "S") || ($dbtype eq "E")) 	# make needed Sybase/msde/Mssql assignments
    {
        open (dbtestsql,">c:\\temp\\emminer.temp.dbtest.sql");
    	print dbtestsql "select \@\@VERSION\n\go\nexit\n";
    	close dbtestsql;

    	system "$sqlcmd -i c:\\temp\\emminer.temp.dbtest.sql -o c:\\temp\\emminer.temp.dbtest.out";
    	if ($debug)
    	   {
    	   	print "debug --- in testdb results of login follow\n";
    	   	system "type c:\\temp\\emminer.temp.dbtest.out";
    	   }    	
    	open (dbtest,"<c:\\temp\\emminer.temp.dbtest.out");
        while (<dbtest>)
           {
           	$failed=index($_,"failed");
           	if ($failed > -1)
           	   {
           	   	   print "Error --> Your DB id, password, or server name caused the login to fail\n\n";
           	   	   print "    id: $emuser\n";
           	   	   print "  pswd: $empswd\n";
           	   	   print "server: $server\n\n";
           	   	   printsyntax();
           	   }
           }  # end of while dbtest
        print "   --> DB access verified\n";
        close dbtest;
    }   
 
  if ($dbtype eq "O")
    {
print "debug in sub testdb, I can tell this is an Oracle db type\n";
print "debug not tested completely, skipping the testdb routine\n";
    }
}  # end of testdb subroutine


#-------------------------------------------
#    sub initvars 
#-----------------------------------

sub initvars ()
{
  if ($debug) {print "--- debug initvars routine\n";}
  gettime();					# capture start time
  $percent="%";
  $nolock=" with (NOLOCK)";			# have set this to null for now  
  $nolock="";
  $emminer_starttime="$hour:$min:$sec";
  $today="$year$mon$mday";
  $u=getlogin;
#  print scout ("emminer.pl run by $u on $mon-$mday at $hour:$min:$sec\n\n");
  use Sys::Hostname;
  $host = hostname;
#  print scout ("---> Machine Hostname: $host\n\n");
  $miscrow=0;					# variable to track what row of the MISC tab we have used  
  $querycnt=1;


  if (($dbtype eq "M") || ($dbtype eq "S") || ($dbtype eq "E")) 	# make needed Sybase/msde/Mssql assignments
    {
      $sqlcmd ="isql -w 9000 -n -U$emuser -P$empass -S$server ";
      $go=" \ngo\n";
      print temp "set nocount on $go";
      $mypat="%";
      $mysubstr="substring";
      $myquote='"';
      $myblank='print " "';
      $mycountq1="'";
      $mycountq2="'";
      $myprint="print";
      $myprintq='"';
      $mysqlpre1="set nocount on  $go";
      $mysqlpre2="";
      $mypat="%";
      $sep=":-:";					# this is the "separator" value between sql columns (helps with parcing)
      $sep1=":-:";	
      $sep1="$myquote$sep1$myquote";
      $sep="$myquote$sep$myquote $mycountq1:-:$mycountq2"; #useful column separator for values returned from sql
    }
  else						# make needed Oracle assignments
    {
      $sqlcmd = "sqlplus -S $emuser/$empass\@$server ";
      $go=";\n";
      print temp "set pagesize 9999;\n";
      print temp "set linesize 9000;\n";
      print temp "set tab off;\n";
      $mypat="*";
      $mysubstr="substr";
      $myquote="'";
      $myblank='prompt';
      $mycountq1='as "';
      $mycountq2='"';
      $myprint="prompt";
      $myprintq="'";
      $mysqlpre1="set pagesize 9999;\n";
      $mysqlpre2="set linesize 400;\n set tab off;\n";    
      $oraspc=" ";
      $mypat="*";  
      $sep=":-:";					# this is the "separator" value between sql columns (helps with parcing)
      $sep1=":-:";	
      $sep1="$myquote$sep1$myquote";
      $sep="$myquote$sep$myquote $mycountq1:-:$mycountq2"; #useful column separator for values returned from sql
      $sep1=$sep;
    }
  if ($dbtype eq "E") 				# adjust for msde support of osql 
    {
      $sqlcmd ="osql -w 9000 -n -U$emuser -P$empass -S$server ";
    }    
  if ($debug) {print "debug --- in initvars the value of sqlcmd=$sqlcmd\n";}
}

#---------------------------------------
# subroutine wrapup
#---------------------------------------

sub wrapup ()
{
   if ($debug) {print "--- debug wrapup routine\n";}
   &namesheet();
   $sheet = $book->Worksheets("Misc");		# set the initial tab the user sees when opening the excel spreadsheet
   $sheet->Activate();
   gettime();					# reaccess the ending time of this routine and put it on the spreadsheet

   $miscrow=$miscrow + 3;     			# put some final data details on the MISC tab
   $sheet->Cells($miscrow,1)->{Value}="   DB Type: $dbtype"; 
   $miscrow=$miscrow + 1;     
   $sheet->Cells($miscrow,1)->{Value}="   EM Version: $emver"; 
   $miscrow=$miscrow + 1;
   $sheet->Cells($miscrow,1)->{Value}="     User: $emuser"; 
   $miscrow=$miscrow + 1;     
   $sheet->Cells($miscrow,1)->{Value}="DB Server: $server";       
   $emminer_endtime="$hour:$min:$sec";
   $miscrow=$miscrow + 2;     
   $sheet->Cells($miscrow,1)->{Value}="Starttime:$emminer_starttime";     
   $miscrow=$miscrow + 1;
   $sheet->Cells($miscrow,1)->{Value}="  Endtime:$emminer_endtime";          
   print ("\n\n   --> Find the EM report in the Excel file c:\\temp\\$excelfile\n");	   

}   # end of wrapup subroutine

#======================     namesheet function     ============================

sub namesheet ()
{

     if ($silent)			     # if silent, use previous runs value	 
       {
         if ($fpref eq "") 
		{
		   $ans="default";
		}
         goto dfltnm; 
       }

     print ("   --> Enter your company name or abbreviation (with no spaces): $fpref");
     $ans = <STDIN>;                          #get options
     chop $ans;                               #remove carrage return
     $ans = lc ($ans);                        #lower case given option


dfltnm:
     if ("$ans" eq "" ) { $ans=$fpref; }
     $fpref=$ans;
     &updconfig();

#     close scout;

     if ($ans ne "")
        {
#         system ("move c:\\temp\\emminer.report c:\\temp\\$ans.emminer.rpt.$year.$mon.$mday.$hour.$min.txt");
#         $rptname="$ans.emminer.rpt.$year.$mon.$mday.$hour.$min.txt";
         $excelfile="$ans.emminer.rpt.$year.$mon.$mday.$hour.$min.xls";
        }
     if ($ans eq "")
        {
#         system ("move c:\\temp\\emminer.report c:\\temp\\emminer.rpt.$year.$mon.$mday.$hour.$min.txt");
#         $rptname="emminer.rpt.$year.$mon.$mday.$hour.$min.txt";
         $excelfile="$emminer.emminer.rpt.$year.$mon.$mday.$hour.$min.xls";
        }
          
cont1:
}

#----------------------------
# subroutine printsyntax
#----------------------------

sub printsyntax()
{
   print "\n         The db server value needed now is to be used in the SQL queries.  So if your not sure if this\n";
   print "         is the instance name, gui server name, tnsnames or interfaces file name, here is how to check.\n";
   print "\n";
   print "         Open a MSDos prompt window and try the command with one of the following syntax (based on DB type).";
   print "\n           For Sybase,MSsql DB test with -->  isql -U<valid EM user id> -P<valid pswd for that id> -S<dbserver name> \n";
   print "\n           For MSDE test with --> osql -U<valid EM user id> -P<valid pswd for that id> -S<dbserver name>\n";
   print "                     Oracle Db test with -->  sqlplus <valid EM user id>/<valid pswd for that id>@<db server name> \n\n";
   print "         If you get a response like SQL>       , you can then exit from SQL and you have verified the id/pswd/and db server name.\n";
   print "         If the login fails, then try a different combination or server name, then use that name for this value.\n";
}