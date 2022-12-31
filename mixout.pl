#!/usr/bin/perl -w
# $DB::single=2;   # remember debug breakpoint
use strict;
use Data::Dumper;               # debug only
no warnings 'recursion';
$DB::deep = 5000;
#$Data::Dumper::Maxrecurse=0;
$Data::Dumper::Maxdepth=8;
my $x;

my $tester = 1;
my $savefile = 0;

my %tax;
my $ta_ref;
my %ttx;

my %case_rcx;

use Win32::OLE;
use Win32::OLE::Const;
use Win32::OLE::Const 'Microsoft Outlook';
use Archive::Zip qw( :ERROR_CODES :CONSTANTS );
use IO::File;
use Win32::OLE::Variant;

use Excel::Writer::XLSX;

my $outlook;

my $aoapath = "C:\\projects\\Packages\\datahealth\\AOAs";

my $gVersion = "0.50000";
my $opt_item;
my $opt_flow;
my $opt_ym;
my $opt_save;
my $opt_max;
my $item_curr = 0;

while (@ARGV) {
   if ($ARGV[0] eq "-h") {
      &GiveHelp;                        # print help and exit
   }
   if ($ARGV[0] eq "-item") {
      shift(@ARGV);
      $opt_item = $ARGV[0] if defined $ARGV[0];
      shift(@ARGV);
   } elsif ($ARGV[0] eq "-ym") {
      shift(@ARGV);
      $opt_ym = $ARGV[0] if defined $ARGV[0];
      shift(@ARGV);
   } elsif ($ARGV[0] eq "-max") {
      shift(@ARGV);
      $opt_max = $ARGV[0] if defined $ARGV[0];
      shift(@ARGV);
   } elsif ($ARGV[0] eq "-save") {
      shift(@ARGV);
      $opt_save = 1;
   } elsif ($ARGV[0] eq "-flow") {
      shift(@ARGV);
      $opt_flow = 1;
   } else {
      die "unknown parameter $ARGV[0]";
   }
}

$opt_item = 10 if !defined $opt_item;
$opt_ym = "2" if !defined $opt_ym;
$opt_max = 20 if !defined $opt_max;
$opt_save = 0 if !defined $opt_save;
$opt_flow = 0 if !defined $opt_flow;
my $otime;



eval {
                $outlook = Win32::OLE->GetActiveObject('Outlook.Application')
};
if ($@ || !defined($outlook)) {
                $outlook = Win32::OLE->new('Outlook.Application', sub{$_[0]->Quit;})
                                or die("Cannot create outlook\n");
}


my $session = $outlook->Session;
my $namespace = $outlook->GetNamespace('MAPI');





# get the proper pst file
my $folder = $namespace->{'Folders'}{'interesting'};

# $Folder->Items->Sort ("[ReceivedTime]"); # Ascending by date

#$folder->Items->Sort ("ReceivedTime"); # Ascending by date


# get all the messages
my $msgs = $folder->{Items};
#my $msgs = $folder->{Items}->Sort("[ReceivedTime]"); # Ascending by date
#my $msgs = $folder->{Items}->Sort("CreationTime"); # Ascending by date
my $ncon = $msgs->{Count};
  print "number of contacts - $ncon\n\n"; # this works


my $con;
my $fn;
my $obase = "C:\\temp\\";
my $po = {};
my $isubject;
my $ibody;
my $ifilename;
my $fn_ct;
my $out;
my $outdir = "C:\\TEMP\\CASE\\";
my $at;
my $odir;
my $oadv_ct;
my $oadv_hi;
my $otems;
my $ohub;
my $oagt;
my $ohost;
my $ocase;
my $omnt;
my $onode;
my $oagts;
my $ofto;
my $oagt_lvl;
my $ohostname;
my $oinstall;
my $ogskit;
my $odur;
my $olevel;
my $omaint;
my $ocopy;
my $oelap;
my $orate;
my $oconfirm;
my $oresult;
my $oworry;
my $odelay;
my $oreplvl;
my $ocat;
my $oca2;
my $oca3;
my $oca4;
my $oca5;
my $oca6;
my $oca7;
my $oca8;
my $oca9;
my $odate;

my $rest;

my %recordx;
my $record_ref;
my $key;

my $i;

foreach my $ii (1 .. $ncon) {
  $out = "";
  $con = $msgs->Item($ii);
  my $isender = $con->SenderName;
  next if substr($isender,0,4) ne "aoa_";
  my $rtime = $con->{ReceivedTime}->Date('yyyyMMdd') . $con->{ReceivedTime}->Time('hhmmss');
  print STDERR "working on $ii $rtime\n" if $opt_flow == 1;
  if ($opt_ym ne "") {
      next if substr($rtime,0,6) ne $opt_ym;
  }
  $item_curr += 1;
  last if $item_curr > $opt_item;
  print STDERR "working on $ii\n" if $opt_flow == 1;
  $isubject = $con->Subject;
  $isubject =~ s/\s+$//;   #trim trailing whitespace
# print $ii . " " . $isubject . "\n";
#$DB::single=2 if $ii == 134;


  if ($isubject eq "testing-testing TEMS Audit") {
     $ibody = $con->Body;
     # 20220207195730 AUDIT 7 /ecurep/sf/TS008/327/TS008327636/2022-02-07/pdcollect-cgslvibmramp003.tar.Z_unpack/
     # TEMSAUDIT 101 27 REMOTE_cgslvibmramp003 *REMOTE 630 tms630fp7:d6350 0 354 CGSLVIBMRAMP003,06.30.07.15,08.00.55.17,(32) https://ibm.biz/BdRHWx
     # 20211217073416 AUDIT 1 /ecurep/sf/TS007/746/TS007746067/2021-12-17/20211217_pdcollect-ttivmon01.tar.Z_unpack/
     # TEMSAUDIT 101 8 HUB_ttivmon01 *LOCAL 0  0 0 ttivmon01,06.30.07.00,08.00.50.69,(32) https://ibm.biz/BdRHWx
     #extract needed data including case number.
     $ibody =~ /(\d+) AUDIT (\d+) (.*?)\nTEMSAUDIT (\d+) (\d+)(.*)/;
     #            1          2      3                4     5     6
     #
     next if !defined $4;    # some cases where second line missing
     $otime = $1;
     $odur = $2;
     $odir = $3;
     $oadv_hi = $4;
     $oadv_ct = $5;
     $rest = $6;
     next if !defined $rest;
     $odir =~ /\/(\S+)\/(\S+)\/(\S+)\/(\S+)\/(\S+)\/(\S+)\/(\S+)\//;    # /ecurep/sf/TS008/327/TS008327636/2022-02-07/pdcollect-cgslvibmramp003.tar.Z_unpack/
     $ocase = $5;
     # Handle cases where nodeid and TEMS type are missing.
     if (substr($rest,0,3) eq "   "){
        $otems = "";
        $ohub = "";
        $rest = substr($rest,3);
     } elsif (substr($rest,0,2) eq "  "){
        $otems = "";
        $rest =~ /(\S+)(.*)/;
        $otems = "";
        $ohub = $1;
        $rest = $2;
     } else {
        $rest =~ / (\S+) (\S+)(.*)/;
        $otems = $1;
        $ohub = $2;
        $rest = $3;
     }


     $rest =~ /(\d+) (.*)/;
     $olevel = $1;
     $rest = $2;
     if (substr($rest,0,1) eq " ") {
        $omaint = "";
        $rest = substr($rest,1);
     } else {
        $rest =~ /(\S+) (.*)/;
        $omaint = $1;
        $rest = $2;
     }
     $rest =~ /(\d+) (\d+) (.*?)\(/;

     $orate = $1;
     $oagt = $2;
     $ohost = $3;

     $ohostname = "";
     $oinstall = "";
     $ogskit = "";
     if ($ohost ne "") {
        $ohost =~  /(.*?)\,(.*?)\,(.*)\,(.*)/;
        $ohostname = $1;
        $oinstall = $2;
        $ogskit = $3;
     }


     $out = "";
     $at = $con->Attachments;
     $fn_ct = $at->Count;
     if ($fn_ct > 0) {
        for ($i=1;$i<=$fn_ct;$i++) {
           $ifilename = $at->Item($i)->{FileName}; # ecureptmpITM_HealthHPxQOGdixQTS008327636_2022-02-07_pdcollect-cgslvibmramp003.tar.Z_unpack.temsaud.csv
           next if index($ifilename,$ocase) == -1;
           $ifilename =~ /$ocase(.*)/;
           $out = $outdir . $ocase . $1;
           last;
        }
     }
     $oreplvl = "N/A";
     if ($out ne "") {
        $at->Item($i)->SaveAsFile($out);
        if (substr($out,-4) eq ".csv") {
           # TEMS Audit report v2.34
           open( FILE, "< $out" ) or die "Cannot open file $out : $!";
           my $inline = <FILE>;
           close FILE;
           $oreplvl = "N/A";
           $inline =~ /TEMS Audit report v(.*)/;
           my $lvl = $1;
           $lvl .= "000" if defined $lvl;
           $oreplvl = $lvl if defined $lvl;
           unlink $out if $opt_save == 0;
        } else {
           warn "item $ii attached file not .csv";
        }

     }


     $key = "TA";
     $record_ref = $recordx{$key};
     if (!defined $record_ref) {
        my %recordref = ( items => {},
                        );
        $recordx{$key} = \%recordref;
        $record_ref = \%recordref;
     }
     my $item_key = $otime . "!" . $ii;
     my $item_ref = $record_ref->{items}{$item_key};
     if (!defined $item_ref) {
        my %itemref =   ( line => $ii,
                          replvl => $oreplvl,
                          start => $otime,
                          dur => $odur,
                          case => $ocase,
                          adv_hi => $oadv_hi,
                          adv_ct => $oadv_ct,
                          maint => $omaint,
                          tems => $otems,
                          hub => $ohub,
                          agt =>$oagt,
                          hostname => $ohostname,
                          install => $oinstall,
                          gskit => $ogskit,
                          odir => $odir,
                        );
        $record_ref->{items}{$item_key} = \%itemref;
        $item_ref = \%itemref;
     }
     if ($ohub eq "*LOCAL") {
        my $kodir = $odir;
        $kodir =~ s/\,/\_/g;
        $ta_ref = $tax{$kodir};
        if (!defined $ta_ref) {
           my %taref =   ( line => $ii,
                           start => $otime,
                           case  => $ocase,
                           tems => $otems,
                         );
           $tax{$kodir} = \%taref;
           $ta_ref = \%taref;
         }
      }


  } elsif ($isubject eq "testing-testing TEMS Audit Log Only") {
     $out = "";
     $at = $con->Attachments;
     $fn_ct = $at->Count;
     if ($fn_ct > 0) {
        for ($i=1;$i<=$fn_ct;$i++) {
           $ifilename = $at->Item($i)->{FileName}; # ecureptmpITM_Health1n86IuNagrTS005333985_2022-02-07_Logs_2022-02-04.part01.rar_unpack.datahealth.csv
           $out = $outdir . $ifilename;
           last;
        }
        $at->Item($i)->SaveAsFile($out) if $out ne "";

        # 2022-02-08 11:07:43 0 logpath /ecurep/log/aoa_ITM_Health/
        # 2022-02-08 11:07:43 0 tmpbase /ecurep/tmp/ITM_Health
        # 2022-02-08 11:07:43 0 tmpclean 0
        # 2022-02-08 11:07:43 0 email jalvord@us.ibm.com
        # 2022-02-08 11:07:43 0 pmr updating 1
        # 2022-02-08 11:07:43 0 start time 20220208110743
        # 2022-02-08 11:07:43 0 temp directory /ecurep/tmp/ITM_Health/f_gwBRIlgt
        # 2022-02-08 11:07:43 0 pdcollect directory /ecurep/sf/TS008/250/TS008250266/2022-02-08/pdcollect-iedu0lvtapp001-DUP0002.tar.Z_unpack

        open( FILE, "< $out" ) or die "Cannot open file $out : $!";
        my @ips = <FILE>;
        close FILE;

        # log file scraping
        my $l = 0;
        foreach my $oneline (@ips)
        {
           $l++;
           chomp($oneline);
           $oneline = substr($oneline,22);
           $oneline =~ /(\S+) (.*)/;
           my $logkey = $1;
           my $rest = $2;
           if ($logkey eq "start") {
              $rest =~ /time (\d+)/;
              $otime = $1;
           } elsif ($logkey eq "pdcollect") {
              $rest =~ /directory (.*)/;
              $odir = $1;
              $odir =~ /\/(\S+)\/(\S+)\/(\S+)\/(\S+)\/(\S+)\/(\S+)\/(\S+)/;    # /ecurep/sf/TS008/327/TS008327636/2022-02-07/pdcollect-cgslvibmramp003.tar.Z_unpack
              $ocase = $5;
              $odir .= '/';
              last;
           }
        }
        unlink $out if $opt_save == 0;

        $key = "AO";
        $record_ref = $recordx{$key};
        if (!defined $record_ref) {
           my %recordref = ( items => {},
                           );
           $recordx{$key} = \%recordref;
           $record_ref = \%recordref;
        }
        my $item_key = $otime . "!" . $ii;
        my $item_ref = $record_ref->{items}{$item_key};
        if (!defined $item_ref) {
           my %itemref =   ( line => $ii,
                             case => $ocase,
                             odir => $odir,
                             start => $otime,
                           );
           $record_ref->{items}{$item_key} = \%itemref;
           $item_ref = \%itemref;
        }
     }

  } elsif ($isubject eq "testing-testing") {
     $ibody = $con->Body;
     # 20220205055941 REFIC 12 /ecurep/sf/TS008/199/TS008199668/2022-02-05/pdcollect-EMS-TEMS-DUP0001.jar_unpack/ REFIC 95 525 29.10% 06.30.06 HTEMS 463 FTO[] https://ibm.biz/BdFrJL
     # 20221026081932 REFIC 5 /ecurep/sf/TS010/995/TS010995232/2022-10-26/pdcollect-NOSMAP4.jar_unpack/ REFIC 105 16 0%   0 FTO[] https://ibm.biz/BdFrJL
     $ibody =~ /(\d+) REFIC (\d+) (.*?) REFIC (\d+) (\d+) (.*%) (.*)/;
     #            1           2     3           4     5     6    7
     $otime = $1;
     $odur = $2;
     $odir = $3;
     $oadv_hi = $4;
     $oadv_ct = $5;
     $oagt_lvl = $6;
     $rest = $7;
     next if !defined $rest;
     if (substr($rest,0,1) eq " ") {
       $omnt = "";
       $rest = substr($rest,1);
     } else {
       $rest =~ /([0-9\.]+) (.*)/;
       $omnt = $1;
       $rest = $2;
     }
     if (substr($rest,0,1) eq " ") {
       $onode = "";
       $rest = substr($rest,1);
     } else {
       $rest =~ /(\S+) (.*)/;
       $omnt = $1;
       $rest = $2;
     }
     $rest =~ /(\d+) FTO(.*?) https/;
     $oagts = $1;
     $ofto = $2;
     if ($ofto eq "[]") {
        $ofto = "N";
     } else {
        $ofto =~ /\[(\S+)]/;
        $ofto = $1;
     }
     $ofto = "N" if $ofto eq "";
     $odir =~ /\/(\S+)\/(\S+)\/(\S+)\/(\S+)\/(\S+)\/(\S+)\/(\S+)\//;    # /ecurep/sf/TS008/327/TS008327636/2022-02-07/pdcollect-cgslvibmramp003.tar.Z_unpack/
     $ocase = $5;
     $out = "";
     $at = $con->Attachments;
     $fn_ct = $at->Count;
     if ($fn_ct > 0) {
        for ($i=1;$i<=$fn_ct;$i++) {
           $ifilename = $at->Item($i)->{FileName}; # ecureptmpITM_Health1n86IuNagrTS005333985_2022-02-07_Logs_2022-02-04.part01.rar_unpack.datahealth.csv
           next if index($ifilename,$ocase) == -1;
           $ifilename =~ /$ocase(.*)/;
           $out = $outdir . $ocase . $1;
           last;
        }
     }
     $oreplvl = "N/A";
     if ($out ne "") {
        $at->Item($i)->SaveAsFile($out);
        if (substr($out,-4) eq ".csv") {
           # ITM Database Health Report 1.79000
           open( FILE, "< $out" ) or die "Cannot open file $out : $!";
           my $inline = <FILE>;
           close FILE;
           $oreplvl = "N/A";
           $inline =~ /ITM Database Health Report (.*)/;
           $oreplvl = $1 if defined $1;
           unlink $out if $opt_save == 0;
        } else {
           warn "item $ii attached file not .csv";
        }
     }


     $key = "TT";
     $record_ref = $recordx{$key};
     if (!defined $record_ref) {
        my %recordref = ( items => {},
                        );
        $recordx{$key} = \%recordref;
        $record_ref = \%recordref;
     }
     my $item_key = $otime . "!" . $ii;
     my $item_ref = $record_ref->{items}{$item_key};
     if (!defined $item_ref) {
        my %itemref =   ( line => $ii,
                          replvl => $oreplvl,
                          start => $otime,
                          dur => $odur,
                          case => $ocase,
                          adv_hi => $oadv_hi,
                          adv_ct => $oadv_ct,
                          agt_lvl => $oagt_lvl,
                          mnt => $omnt,
                          tems => $onode,
                          agts => $oagts,
                          fto => $ofto,
                          odir => $odir,
                        );
        $record_ref->{items}{$item_key} = \%itemref;
        $item_ref = \%itemref;
      }
      my $kodir = $odir;
      $kodir =~ s/\,/\_/g;
      $ttx{$kodir} = $ii;
  } elsif ($isubject eq "event-audit") {
     $ibody = $con->Body;
     # EVENTAUDIT 65 3 604800 seconds 6027 events[1.49/min] Confirm[99.85%] results[1.35K/min] worry[0.27%] delay[0.00] copy[604800,1.49,99.85%,1.35K,0.27%,0.00]
     # 20220205060000 REFIC 4 /ecurep/sf/TS008/199/TS008199668/2022-02-05/pdcollect-EMS-TEMS-DUP0001.jar_unpack/ EVENTAUDIT 65 3 604800 seconds 6027 events[1.49/min] Confirm[99.85%] results[1.35K/min] worry[0.27%] delay[0.00] copy[604800,1.49,99.85%,1.35K,0.27%,0.00]
     $ibody =~ /EVENTAUDIT (\d+) (\d+) .*copy.*\n(\d+) REFIC (\d+) (.*?) EVENTAUDIT .*?copy\[(.*?)]/;
     $otime = $3;
     $odur = $4;
     $odir = $5;
     $oadv_hi = $1;
     $oadv_ct = $2;
     $ocopy =  $6;
     next if !defined $odir;
     $odir =~ /\/(\S+)\/(\S+)\/(\S+)\/(\S+)\/(\S+)\/(\S+)\/(\S+)\//;    # /ecurep/sf/TS008/327/TS008327636/2022-02-07/pdcollect-cgslvibmramp003.tar.Z_unpack/
     $ocase = $5;

     $ocopy =~ /(\d+)\,(\d+\.?\d*),(\d+\.?\d*%),(\d+\.?\d*K),(\d+\.?\d*%),(\d+\.?\d*)/;
     $oelap = $1;
     $orate = $2;
     $oconfirm = $3;
     $oresult = $4;
     $oworry = $5;
     $odelay = $6;

     $out = "";
     $at = $con->Attachments;
     $fn_ct = $at->Count;
     if ($fn_ct > 0) {
        for ($i=1;$i<=$fn_ct;$i++) {
           $ifilename = $at->Item($i)->{FileName}; # ecureptmpITM_Health1n86IuNagrTS005333985_2022-02-07_Logs_2022-02-04.part01.rar_unpack.datahealth.csv
           next if index($ifilename,$ocase) == -1;
           $ifilename =~ /$ocase(.*)/;
           $out = $outdir . $ocase . $1;
           last;
        }
     }
     $oreplvl = "N/A";
     if ($out ne "") {
        $at->Item($i)->SaveAsFile($out);
        if (substr($out,-4) eq ".csv") {
           # Situation Status History Audit Report 1.39000
           open( FILE, "< $out" ) or die "Cannot open file $out : $!";
           my $inline = <FILE>;
           close FILE;
           $oreplvl = "N/A";
           $inline =~ /Situation Status History Audit Report (.*)/;
           $oreplvl = $1 if defined $1;
           unlink $out if $opt_save == 0;
        } else {
           warn "item $ii attached file not .csv";
        }
     }

     $key = "EV";
     $record_ref = $recordx{$key};
     if (!defined $record_ref) {
        my %recordref = ( items => {},
                        );
        $recordx{$key} = \%recordref;
        $record_ref = \%recordref;
     }
     my $item_key = $otime . "!" . $ii;
     my $item_ref = $record_ref->{items}{$item_key};
     if (!defined $item_ref) {
        my %itemref =   ( line => $ii,
                          replvl => $oreplvl,
                          start => $otime,
                          odir => $odir,
                          dur => $odur,
                          case => $ocase,
                          adv_hi => $oadv_hi,
                          adv_ct => $oadv_ct,
                          elap => $oelap,
                          rate => $orate,
                          confirm => $oconfirm,
                          result => $oresult,
                          worry => $oworry,
                          delay => $odelay,
                        );
        $record_ref->{items}{$item_key} = \%itemref;
        $item_ref = \%itemref;
     }
  } elsif (index($isubject,"AOA Critical Alert") >=0) {    # re: TS005938651 AOA Critical Alert
     $ibody = $con->Body;
     # ITM6 Analysis On Arrival Critical Error Report /ecurep/sf/TS008/327/TS008327636/2022-02-07/pdcollect cgslvibmramp003.tar.Z_unpack
     # temsaud.crit:Hub/Remote sync[2] Max/secs[0] SynDrq fails[0] http://ibm.biz/BdYVYG
     # temsaud.crit:Reconnection from remote TEMS to hub TEMS - 2 times http://ibm.biz/BdYVYG
     # temsaud.crit:TEMS has lost connection to HUB 3 times http://ibm.biz/BdYVYG
     # copy[3,3,0,0,0,0,0,0,0,]

     $ibody =~ /Report (.*?)\n/s;
     $odir = $1;
     chop $odir;
     $odir .= '/';
     $ibody =~ /copy\[(.*?)\]/s;
     $ocopy = $1;
     $ocopy =~ /(\d+),(\d+),(\d+),(\d+),(\d+),(\d+),(\d+),(\d+),(\d+),/;
     $ocat = $1;
     $oca2 = $2;
     $oca3 = $3;
     $oca4 = $4;
     $oca5 = $5;
     $oca6 = $6;
     $oca7 = $7;
     $oca8 = $8;
     $oca9 = $9;

     $odir =~ /\/(\S+)\/(\S+)\/(\S+)\/(\S+)\/(\S+)\/(\S+)\/(\S+)\//;    # /ecurep/sf/TS008/327/TS008327636/2022-02-07/pdcollect-cgslvibmramp003.tar.Z_unpack/
     $ocase = $5;

     my @words = split /\n/,$ibody;
     my @lines_c;
     foreach my $li (@words) {
        last if substr($li,0,4) eq "copy";
        next if index($li,":") == -1;
        push @lines_c,$li;
     }
     #s/\\/\//g;
     if ($#lines_c >= 0) {
        # ts006671383_2021-09-01_pdcollect-momtemsd-dup0001.tar.z_unpack.eventaud.csv
        $odir =~ /$ocase(.*)/;
        my $ifn = $1;
        chop $ifn;
        $ifn =~ s/\//_/g;
        my $fn = $outdir . $ocase . "_" . $ifn . ".aoaci.txt";
        open( FILE, "+> $fn" ) or die "Cannot open file $fn : $!";
        foreach my $oline (@lines_c) {
            print FILE $oline;
        }
        close FILE;
     }



     $otime = $con->{ReceivedTime}->Date('yyyyMMdd') . $con->{ReceivedTime}->Time('hhmmss');
     $odate = $con->{ReceivedTime}->Date();


     $key = "CA";
     $record_ref = $recordx{$key};
     if (!defined $record_ref) {
        my %recordref = ( items => {},
                        );
        $recordx{$key} = \%recordref;
        $record_ref = \%recordref;
     }
     my $item_key = $otime . "!" . $ii;
     my $item_ref = $record_ref->{items}{$item_key};
     if (!defined $item_ref) {
        my %itemref =   ( line => $ii,
                          case => $ocase,
                          odir => $odir,
                          cat => $ocat,
                          ca2 => $oca2,
                          ca3 => $oca3,
                          ca4 => $oca4,
                          ca5 => $oca5,
                          ca6 => $oca6,
                          ca7 => $oca7,
                          ca8 => $oca8,
                          ca9 => $oca9,
                          start => $otime,
                          date => $odate,
                        );
        $record_ref->{items}{$item_key} = \%itemref;
        $item_ref = \%itemref;
     }
     next;
  } elsif (index($isubject,"itm_tep_recert") >=0) {  # itm_tep_recert 75724 TS008346468 Success
    $isubject =~ /itm_tep_recert \d+ (\S+) (\S+)/;
    $ocase = $1;
    $odate = $con->{ReceivedTime}->Date();
    $otime = $con->{ReceivedTime}->Date('yyyyMMdd') . $con->{ReceivedTime}->Time('hhmmss');
    $key = "RC";
     $record_ref = $recordx{$key};
     if (!defined $record_ref) {
        my %recordref = ( items => {},
                        );
        $recordx{$key} = \%recordref;
        $record_ref = \%recordref;
     }
     my $item_key = $otime . "!" . $ii;
     my $item_ref = $record_ref->{items}{$item_key};
     if (!defined $item_ref) {
        my %itemref =   ( line => $ii,
                          case => $ocase,
                          date => $odate,
                          time => $otime,
                        );
        $record_ref->{items}{$item_key} = \%itemref;
        $item_ref = \%itemref;
     }
  } else {
    warn "Unknown subject $isubject";
  }
}

foreach my $f (keys %tax) {
   next if defined $ttx{$f};
   $ta_ref = $tax{$f};
   $key = "MT";
   $record_ref = $recordx{$key};
   if (!defined $record_ref) {
      my %recordref = ( items => {},
                      );
      $recordx{$key} = \%recordref;
      $record_ref = \%recordref;
   }
   my $item_key = $ta_ref->{start} . "!" . $ta_ref->{line};
   my $item_ref = $record_ref->{items}{$item_key};
   if (!defined $item_ref) {
      my %itemref =   ( line => $ta_ref->{line},
                        case => $ta_ref->{case},
                        start => $ta_ref->{start},
                        tems => $ta_ref->{tems},
                        odir => $f,
                      );
      $record_ref->{items}{$item_key} = \%itemref;
      $item_ref = \%itemref;
      my $x = 1;
     }
}


# At this point the %recordx hash array has a number of records ready to store
# in a work XLS file ready for transfer to a history worksheet.

my $outxls = "C:\\TEMP\\stage.xlsx";
unlink $outxls;

my $xl = Win32::OLE->new('Excel.Application');
$xl->{EnableEvents} = 0;
$xl->{Visible} = 0;

my $wb = $xl->Workbooks->Add;                                                  # add three worksheets
$wb->Worksheets->Add({After => $wb->Worksheets($wb->Worksheets->{Count})});    # add three more worksheets
$wb->Worksheets->Add({After => $wb->Worksheets($wb->Worksheets->{Count})});    # add three more worksheets
$wb->Worksheets->Add({After => $wb->Worksheets($wb->Worksheets->{Count})});    # add three more worksheets
my $sht_rc = $wb->Sheets(1);      # Recerts
my $sht_ta = $wb->Sheets(2);      # TEMS Audit Reports
my $sht_tt = $wb->Sheets(3);      # Database Health Checker
my $sht_ev = $wb->Sheets(4);      # Event Audit
my $sht_ca = $wb->Sheets(5);      # Event Audit
my $sht_sf = $wb->Sheets(6);      # S/F cases
$sht_rc->{Name} = 'Resigns';
$sht_ta->{Name} = 'Audit';
$sht_tt->{Name} = 'AOAs';
$sht_ev->{Name} = 'Events';
$sht_ca->{Name} = 'AOACI';
$sht_sf->{Name} = 'SalesForce';

#my $number_format = $wb->add_format(
#    num_format => '0',
#);
#
#$sht_ta->set_column( 'D:D', undef, $number_format);


foreach my $r ( sort { $a cmp $b } keys %recordx) {
   $record_ref = $recordx{$r};
   if ($r eq "RC") {
      no strict;
     foreach my $i ( sort {$record_ref->{items}{$a}->{time} cmp $record_ref->{items}{$b}->{time} } keys %{$record_ref->{items}} ) {
        my $item_ref =  $record_ref->{items}{$i};
        #{
        #  'time' => '20220207114016',
        #  'line' => 36,
        #  'date' => '2/7/2022',
        #  'case' => 'TS008346468'
        #};
        $sht_rc->Cells(1,1)->EntireRow->Insert;
        $sht_rc->Cells(1,1)->{Value} = $item_ref->{case};
        $sht_rc->Cells(1,6)->{Value} = $item_ref->{date};
        $sht_rc->Cells(1,11)->{Value} = 1;
        $case_rcx{$item_ref->{case}} .= ",RC";
     }
     use strict;
   } elsif ($r eq "TA") {
      no strict;
     foreach my $i ( sort {$record_ref->{items}{$a}->{start} cmp $record_ref->{items}{$b}->{start} } keys %{$record_ref->{items}} ) {
        my $item_ref =  $record_ref->{items}{$i};
        #{
        #   'replvl' => '2.34',
        #   'start' => '20220207195730',
        #   'dur' => '7',
        #   'case' => 'TS008327636',
        #   'adv_hi' => '101',
        #   'adv_ct' => '27',
        #   'maint' => 'tms630fp7:d6350',
        #   'tems' => 'REMOTE_cgslvibmramp003',
        #   'hub' => '*REMOTE',
        #   'agt' => '354',
        #   'hostname' => 'CGSLVIBMRAMP003',
        #   'install' => '06.30.07.15',
        #   'gskit' => '08.00.55.17',
        #   'line' => 3,
        #   'odir' => '/ecurep/sf/TS008/327/TS008327636/2022-02-07/pdcollect-cgslvibmramp003.tar.Z_unpack/
        #}
        $sht_ta->Cells(1,1)->EntireRow->Insert;
        $sht_ta->Cells(1,1)->{Value} = $item_ref->{replvl};
        $sht_ta->Cells(1,4)->{Value} = $item_ref->{start};
        $sht_ta->Range("D1:D1")->{NumberFormat} = "0";
        $sht_ta->Cells(1,5)->{Value} = $item_ref->{dur};
        $sht_ta->Cells(1,6)->{Value} = $item_ref->{case};
        $sht_ta->Cells(1,7)->{Value} = $item_ref->{adv_hi};
        $sht_ta->Cells(1,8)->{Value} = $item_ref->{adv_ct};
        $sht_ta->Cells(1,17)->{Value} = $item_ref->{maint};
        $sht_ta->Cells(1,18)->{Value} = $item_ref->{tems};
        $sht_ta->Cells(1,19)->{Value} = $item_ref->{hub};
        $sht_ta->Cells(1,20)->{Value} = $item_ref->{agt};
        $sht_ta->Cells(1,21)->{Value} = $item_ref->{hostname};
        $sht_ta->Cells(1,22)->{Value} = $item_ref->{install};
        $sht_ta->Cells(1,23)->{Value} = $item_ref->{gskit};
        $sht_ta->Cells(1,27)->{Value} = $item_ref->{odir};
        $sht_ta->Cells(1,32)->{Value} = $item_ref->{line};
        $case_rcx{$item_ref->{case}} .= ",RA";
      }
      use strict;
   } elsif ($r eq "AO") {
      no strict;
     foreach my $i ( sort {$record_ref->{items}{$a}->{start} cmp $record_ref->{items}{$b}->{start} } keys %{$record_ref->{items}} ) {
        my $item_ref =  $record_ref->{items}{$i};
        #{
        #   'line' => 49,
        #   'start' => '20220208110743',
        #   'case' => 'TS008250266',
        #   'odir' => '/ecurep/sf/TS008/250/TS008250266/2022-02-08/pdcollect-iedu0lvtapp001-DUP0002.tar.Z_unpack/'
        #}
        $sht_ta->Cells(1,1)->EntireRow->Insert;
        $sht_ta->Cells(1,4)->{Value} = $item_ref->{start};
        $sht_ta->Range("D1:D1")->{NumberFormat} = "0";
        $sht_ta->Cells(1,6)->{Value} = $item_ref->{case};
        $sht_ta->Cells(1,27)->{Value} = $item_ref->{odir};
        $sht_ta->Cells(1,32)->{Value} = $item_ref->{line};
#       $sht_tt->Cells(1,1)->EntireRow->Insert;
#       $sht_tt->Cells(1,1)->{Value} = "N/A";
#       $sht_tt->Cells(1,4)->{Value} = $item_ref->{start};
#       $sht_tt->Range("D1:D1")->{NumberFormat} = "0";
#       $sht_tt->Cells(1,6)->{Value} = $item_ref->{case};
#       $sht_tt->Cells(1,29)->{Value} = $item_ref->{odir};
#       $sht_tt->Cells(1,34)->{Value} = $item_ref->{line};
        $case_rcx{$item_ref->{case}} .= ",AO";
      }
      use strict;
   } elsif ($r eq "TT") {
      no strict;
     foreach my $i ( sort {$record_ref->{items}{$a}->{start} cmp $record_ref->{items}{$b}->{start} } keys %{$record_ref->{items}} ) {
        my $item_ref =  $record_ref->{items}{$i};
        #{
        #   'replvl' => 'N/A',
        #   'start' => '20220205055941',
        #   'dur' => '12',
        #   'case' => 'TS008199668',
        #   'adv_hi' => '95'
        #   'adv_ct' => '525',
        #   'agt_lvl' => '29.10%',
        #   'mnt' => '06.30.06',
        #   'tems' => 'HTEMS',
        #   'agts' => '463',
        #   'fto' => 'N',
        #   'odir' => '/ecurep/sf/TS008/199/TS008199668/2022-02-05/pdcollect-EMS-TEMS-DUP0001.jar_unpack/',
        #   'line' => 4,
        #}

        $sht_tt->Cells(1,1)->EntireRow->Insert;
        $sht_tt->Cells(1,1)->{Value} = $item_ref->{replvl};
        $sht_tt->Cells(1,4)->{Value} = $item_ref->{start};
        $sht_tt->Range("D1:D1")->{NumberFormat} = "0";
        $sht_tt->Cells(1,5)->{Value} = $item_ref->{dur};
        $sht_tt->Cells(1,6)->{Value} = $item_ref->{case};
        $sht_tt->Cells(1,7)->{Value} = $item_ref->{adv_hi};
        $sht_tt->Cells(1,8)->{Value} = $item_ref->{adv_ct};
        $sht_tt->Cells(1,16)->{Value} = $item_ref->{agt_lvl};
        $sht_tt->Cells(1,17)->{Value} = $item_ref->{mnt};
        $sht_tt->Cells(1,18)->{Value} = $item_ref->{tems};
        $sht_tt->Cells(1,19)->{Value} = $item_ref->{agts};
        $sht_tt->Cells(1,20)->{Value} = $item_ref->{fto};
        $sht_tt->Cells(1,29)->{Value} = $item_ref->{odir};
        $sht_tt->Cells(1,34)->{Value} = $item_ref->{line};
        $case_rcx{$item_ref->{case}} .= ",TT";
      }
      use strict;
   } elsif ($r eq "MT") {
      no strict;
     foreach my $i ( sort {$record_ref->{items}{$a}->{start} cmp $record_ref->{items}{$b}->{start} } keys %{$record_ref->{items}} ) {
        my $item_ref =  $record_ref->{items}{$i};
        #{
        #   'replvl' => 'N/A',
        #   'start' => '20220205055941',
        #   'dur' => '12',
        #   'case' => 'TS008199668',
        #   'adv_hi' => '95'
        #   'adv_ct' => '525',
        #   'agt_lvl' => '29.10%',
        #   'mnt' => '06.30.06',
        #   'tems' => 'HTEMS',
        #   'agts' => '463',
        #   'fto' => 'N',
        #   'odir' => '/ecurep/sf/TS008/199/TS008199668/2022-02-05/pdcollect-EMS-TEMS-DUP0001.jar_unpack/',
        #   'line' => 4,
        #}

        $sht_tt->Cells(1,1)->EntireRow->Insert;
        $sht_tt->Cells(1,4)->{Value} = $item_ref->{start};
        $sht_tt->Range("D1:D1")->{NumberFormat} = "0";
        $sht_tt->Cells(1,6)->{Value} = $item_ref->{case};
        $sht_tt->Cells(1,17)->{Value} = $item_ref->{tems};
        $sht_tt->Cells(1,29)->{Value} = $item_ref->{odir};
        $sht_tt->Cells(1,34)->{Value} = $item_ref->{line};
        $case_rcx{$item_ref->{case}} .= ",MT";
      }
      use strict;
   } elsif ($r eq "EV") {
      no strict;
     foreach my $i ( sort {$record_ref->{items}{$a}->{start} cmp $record_ref->{items}{$b}->{start} } keys %{$record_ref->{items}} ) {
        my $item_ref =  $record_ref->{items}{$i};
        #{
        #   'replvl' => '1.39',
        #   'start' => '20220205060000',
        #   'dur' => '4',
        #   'case' => 'TS008199668',
        #   'adv_hi' => '65',
        #   'adv_ct' => '3',
        #   'elap' => '604800',
        #   'rpte' => '1.49'
        #   'confirm' => '99.85%',
        #   'result' => '1.35K',
        #   'worry' => '0.27%',
        #   'delay' => '0.00',
        #   'odir' => '/ecurep/sf/TS008/199/TS008199668/2022-02-05/pdcollect-EMS-TEMS-DUP0001.jar_unpack/',
        #   'line' => 6,
        #}

        $sht_ev->Cells(1,1)->EntireRow->Insert;
        $sht_ev->Cells(1,1)->{Value} = $item_ref->{replvl};
        $sht_ev->Cells(1,4)->{Value} = $item_ref->{start};
        $sht_ev->Range("D1:D1")->{NumberFormat} = "0";
        $sht_ev->Cells(1,5)->{Value} = $item_ref->{dur};
        $sht_ev->Cells(1,6)->{Value} = $item_ref->{case};
        $sht_ev->Cells(1,7)->{Value} = $item_ref->{adv_hi};
        $sht_ev->Cells(1,8)->{Value} = $item_ref->{adv_ct};
        $sht_ev->Cells(1,17)->{Value} = $item_ref->{elap};
        $sht_ev->Cells(1,18)->{Value} = $item_ref->{rate};
        $sht_ev->Cells(1,19)->{Value} = $item_ref->{confirm};
        $sht_ev->Cells(1,20)->{Value} = $item_ref->{result};
        $sht_ev->Cells(1,21)->{Value} = $item_ref->{worry};
        $sht_ev->Cells(1,22)->{Value} = $item_ref->{delay};
        $sht_ev->Cells(1,23)->{Value} = $item_ref->{odir};
        $sht_ev->Cells(1,28)->{Value} = $item_ref->{line};
        $case_rcx{$item_ref->{case}} .= ",EV";
      }
   } elsif ($r eq "CA") {
      no strict;
     foreach my $i ( sort {$record_ref->{items}{$a}->{start} cmp $record_ref->{items}{$b}->{start} } keys %{$record_ref->{items}} ) {
        my $item_ref =  $record_ref->{items}{$i};
        #{
        # 'case' => 'TS006344412',
        # 'date' => '8/1/2021',
        # 'cat' => '1',
        # 'ca2' => '0',
        # 'ca3' => '1',
        # 'ca4' => '0',
        # 'ca5' => '0',
        # 'ca6' => '0',
        # 'ca7' => '0',
        # 'ca8' => '0'
        # 'ca9' => '0',
        # 'odir' => '/ecurep/sf/TS006/344/TS006344412/2021-08-01/pdcollect-shssapppsys02.tar.Z_unpack/',
        # 'line' => 4466,
        #};

        $sht_ca->Cells(1,1)->EntireRow->Insert;
        $sht_ca->Cells(1,1)->{Value} = $item_ref->{case};
        $sht_ca->Cells(1,6)->{Value} = $item_ref->{date};
        $sht_ca->Cells(1,9)->{Value} = $item_ref->{cat};
        $sht_ca->Cells(1,10)->{Value} = $item_ref->{ca2};
        $sht_ca->Cells(1,11)->{Value} = $item_ref->{ca3};
        $sht_ca->Cells(1,12)->{Value} = $item_ref->{ca4};
        $sht_ca->Cells(1,13)->{Value} = $item_ref->{ca5};
        $sht_ca->Cells(1,14)->{Value} = $item_ref->{ca6};
        $sht_ca->Cells(1,15)->{Value} = $item_ref->{ca7};
        $sht_ca->Cells(1,16)->{Value} = $item_ref->{ca8};
        $sht_ca->Cells(1,17)->{Value} = $item_ref->{ca9};
        $sht_ca->Cells(1,19)->{Value} = $item_ref->{odir};
        $sht_ca->Cells(1,24)->{Value} = $item_ref->{line};
        $case_rcx{$item_ref->{case}} .= ",CA";
      }
      use strict;
   }
}


my $caserc_ct = scalar keys %case_rcx;
if ($caserc_ct > 0) {
     foreach my $c ( sort {$a cmp $b } keys %case_rcx ) {
        $sht_sf->Cells(1,1)->EntireRow->Insert;
        $sht_sf->Cells(1,1)->{Value} = $c;
        $sht_sf->Cells(1,2)->{Value} = $case_rcx{$c};
     }
     my $sumline = "";
     my $sum_ct = 0;
     foreach my $c ( sort {$a cmp $b } keys %case_rcx ) {
        $sum_ct += 1;
        $sumline .= $c;
        if ($sum_ct >= $opt_max) {
           $sht_sf->Cells(1,1)->EntireRow->Insert;
           $sht_sf->Cells(1,1)->{Value} = $sumline;
          $sumline = "";
          $sum_ct = 0
        } else {
           $sumline .= ",";
        }
     }
     if ($sumline ne "") {
        chop $sumline if substr($sumline,-1) eq ",";
        $sht_sf->Cells(1,1)->EntireRow->Insert;
        $sht_sf->Cells(1,1)->{Value} = $sumline;
     }
}



$wb->SaveAs($outxls);
$xl->Quit();


exit 0;

##!/usr/bin/perl
#use strict;
#use warnings;
#
#use Win32::OLE;
#use Win32::OLE::Const 'Microsoft Outlook';
#my $filename = 'c:\\net.txt';
#open( my $output_fh, ">", $filename ) or die $!;
#
#my $outlook = Win32::OLE->new('Outlook.Application')
#    or die "Failed Opening Outlook.";
#
#my $namespace    = $outlook->GetNamespace("MAPI");
#my $archive        = $namespace->GetDefaultFolder(6)->Folders('Archive');
#my $deletedItems = $namespace->GetDefaultFolder(3);
#my $items        = $archive->Items;
#
#foreach my $msg ( $items->in ) {
#    if ( $msg->{Subject} =~ m/^test/ ) {
#        print $msg ->{Subject}, "\n";
#        print {$output_fh} $msg->{Body};
#        $msg->Move($deletedItems);
#    }
#}
#
#close($output_fh);
sub GiveHelp
{
  $0 =~ s|(.*)/([^/]*)|$2|;
  print <<"EndOFHelp";

  $0 v$gVersion

  This script reads outlook emails.

  Default values:
    item = 1

  Run as follows:
    $0  <options> log_file

  Options
    -h              display help information
    -item           number of items

  Examples:
    $0  logfile > results.csv

EndOFHelp
}
exit;
