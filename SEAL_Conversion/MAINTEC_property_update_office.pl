#!/usr/local/bin/seppperl
#!/usr/local/bin/seppperl -d:ptkdb
# ---------------------------------------------------------------
# History:
# $Log: VAI_property_update_office.pl,v $
# Revision 1.7  2006/09/12 09:42:12  ludwig
# Added xls as possible format to be property updated. VAI said, maintec
# components can handle xls and xlt files too.
#
# Revision 1.6  2006/07/27 12:12:38  ludwig
# Fixed bug interpreting PROPERTIES section name.
#
# Revision 1.5  2006/01/05 15:38:05  ludwig
# Fixed Jira-Ticket DPFc-107. Adapt Online-Doku of each WU for Perl 5.8.
# Else the ptk debugger gets problems if the pod doku resides on the beginning
# of the perl file.
#
# Revision 1.4  2005/10/25 13:47:02  ludwig
# Prevent now property update office from dpf4convert standard.
#
# Revision 1.3  2005/10/21 17:09:19  ludwig
# Changed reading of property update file to changed separator (from " "
# to "=").
#
# Revision 1.2  2005/09/30 11:47:15  ludwig
# Only word files can be updated by maintec property update tool. Other file
# types now are ignored.
# Add timeout handling for calls of maintec tools.
#
# Revision 1.1  2005/08/17 10:14:06  ludwig
# Moved from server/dpf/scripts/dpf4convert to here.
#
# Revision 1.3  2005/08/16 16:36:27  ludwig
# *** empty log message ***
#
# Revision 1.2  2005/08/16 16:36:04  ludwig
# CheckValidFormat shall not exit with positive error code, because other
# formats like dwg shall be converted in the whole process further.
#
# Revision 1.1  2005/08/01 10:05:17  ludwig
# This is a customer working unit!
# Maintec now can property update the file without having the document in
# the same directory according to the exe-file.
# Added dpftracker error showing functionality.
# Add check of valid property update parameters according to vai rules.
# Seal message are only added. Maintec than just crashes.
#
#
# ---------------------------------------------------------------
use dpf::wu;
use File::Basename;
use File::Copy;
use seppperl::msghandler;   #- Modul to provide multiple languages independently.

my ($rLang);
$rLang = "seppperl::msghandler"->
                        New($Language, "os", "lib", "gates", "tools", "dpf");
use libconvdpf2;
use strict;
# -------------------------------------------------------
# initialize variables
my $JobDir = $Par{"JOBDIR"};
my $Currfilename = $Par{"CURRFILENAME"};
$Currfilename =~ s/\//\\/g;
my $Currfiletype = $Par{"CURRFILETYPE"};
my $Dpf4cPropertyUpdateEnable = $Par{"DPF4C_PROPERTYUPDATE_ENABLE"};
my $Dpf4cPropertyUpdateProvideFiles =
                                    $Par{"DPF4C_PROPERTYUPDATE_PROVIDE_FILES"};
my $Dpf4cPropertyUpdateFile = $Par{"DPF4C_PROPERTYUPDATE_FILE"};
my $PropertyUpdateExe = $ENV{"PLSTOOLS"} . 
                    "\\vai_tools\\OfficePropertyUpdate\\MAINTECSAP2Word.exe";
my $VaiPropertyUpdateOfficeWrite = $Par{"VAI_PROPERTY_UPDATE_OFFICE_WRITE"};
my $DirMaintecOffice = $ENV{"PLSTOOLS"} . "\\vai_tools\\OfficePropertyUpdate";
my $Dpf4cCommonDestFomat = $Par{"DPF4C_COMMON_DEST_FORMAT"};
my $Dpf4cCommonTimeout = &libconvdpf2::GetDpf4CTimeout 
                                        ($Par{"DPF4C_COMMON_TIMEOUT"});

my %TitleHash = (   "EASY_SUBJECT" => "Subject",
                    "EASY_DESCRIPTION" => "vaiDescription",
                    "EASY_KEYWORDS" => "Keywords",
                    "EASY_VAIASBUILT" => "vaiAsBuilt",
                    "EASY_VAICOMPANY" => "vaiCompany",
                    "EASY_VAIDEPARTMENT" => "vaiDepartment",
                    "EASY_AUTHOR" => "Author",
                    "VAIVERSION" => "vaiRevision",
                    "VAI_STATUS" => "vaiStatus",
                    "SCSNUM" => "vaiSCSNumber",
                    "DTP_TITLE" => "vaiDocType",
                    "DTP" => "vaiDocTypeCode",
                    "VAICORR_NEW_NUMBER" => "vaiCorrespNo",
                    "PR_DISPLAY_TO" => "vaiTo",
                    "PR_CLIENT_SUBMIT_TIME" => "vaiSentReceived",
                    "EASY_TITLE1" => "vaiTitle1",
                    "EASY_PROJECTCODE" => "vaiProjectCode",
                    "EASY_TITLE2" => "vaiTitle2",
                    "EASY_PLANTCODE" => "vaiPlantCode",
                    "EASY_TITLE3" => "vaiTitle3",
                    "EASY_AREA" => "vaiAreaCode",
                    "EASY_TITLE4" => "vaiTitle4",
                    "EASY_EQUIPMENT" => "vaiEquipCode",
                    "SGC_TITLE" => "vaiSGC",
                    "SGC" => "vaiSGCCode",
                    "SEQUNO" => "vaiSeqNo",
                    "EASY_SECLEVEL" => "vaiSecLevel",
                    "EASY_ENGPACKAGE" => "vaiEngPackage",
                    "EASY_SECLEVTXT" => "vaiSecLevelTxt",
                    "EASY_ENGPACKTXT" => "vaiEngPackTxt");
                    
my @SummaryOptions = ("Title", "Subject", "Category", "Keywords",
                      "Template", "Page Count", "Word Count", 
                      "Character Count", "Byte Count", "Lines",
                      "Paragraphs", "Scale", "Links Dirty?", "Comments",
                      "Author", "Last Saved By", "Revision Number", 
                      "Application Name", "Company Name", "Date of Creation",
                      "Date Last Saved", "Last Printed", "Edit time",
                      "MaintecBuilt"); # exception for vai against their rules

my $ErrorText;
if ($Dpf4cPropertyUpdateEnable ne "Y")
    {
    &Log ("No property update to do because of dpf ".
          "parameter DPF4C_PROPERTYUPDATE_ENABLE.", "I");
    exit 0;
    }

&CheckValidFormat ($Currfiletype);

# check if files available
if (!(-e $Currfilename))
    {
    $ErrorText = "Error: File CURRFILENAME <$Currfilename> is not available.";
    &libconvdpf2::SetDpfStatus (\%Par, 1, $ErrorText);
    exit 1;
    }
if (!(-e $Dpf4cPropertyUpdateFile))
    {
    $ErrorText = "Error: Propertyupdate file DPF4C_PROPERTYUPDATE_FILE ".
                 "<$Dpf4cPropertyUpdateFile> is not available.";
    &libconvdpf2::SetDpfStatus (\%Par, 2, $ErrorText);
    exit 2;
    }
if (!(-e $PropertyUpdateExe))
    {
    $ErrorText = "Error: Propertyupdate exe file ".
                 "<$PropertyUpdateExe> is not available.";
    &libconvdpf2::SetDpfStatus (\%Par, 3, $ErrorText);
    exit 3;
    }

# expand $Dpf4cPropertyUpdateFile to complete Path if necessary
my ($file, $path, $ext) = fileparse ($Dpf4cPropertyUpdateFile, "\.[^.]*");
if (($path eq "")||($path =~ /\.[\\\/]/))
    {
    $Dpf4cPropertyUpdateFile = "$JobDir/$Dpf4cPropertyUpdateFile";
    &Log ("Expanding path of property update file to ".
          "<$Dpf4cPropertyUpdateFile>", "D");
    }

# save not property updated office file
if (! copy ($Currfilename, "$Currfilename.propertyupdatesave"))
    {
    $ErrorText = "Error saving original office document ".
                  "'$Currfilename' to '$Currfilename.propertyupdatesave': $!";
    &libconvdpf2::SetDpfStatus (\%Par, 4, $ErrorText);
    exit 4;
    }

# registry thingies
&DoOfficeRegistry ();

# do property update for all properties in PropertyUpdateFile
my $stat;
my %UpdateProperties = &GetPropertiesAndMapVai ($Dpf4cPropertyUpdateFile);
my $UpdateProperty;
foreach $UpdateProperty (sort keys %UpdateProperties)
    {
    my $PropertyValue = $UpdateProperties{$UpdateProperty};
    &Log ("Start property update for word entry <$UpdateProperty> with ".
          "value <$PropertyValue>.", "I");
    my $CheckErrorText = "";
    ## 2007-03-14 - No further check of vai-property nesessary
    # $CheckErrorText = &CheckPropertyKey ($UpdateProperty, $PropertyValue);
    my $IsCustomProperty = "n";      
    if ($UpdateProperty =~ /^vai/)
        { $IsCustomProperty = "y"; }
    my @cmd = ("$PropertyUpdateExe",   
            "w", # means to use WriteProp
            "\"$Currfilename\"",
            "$UpdateProperty",
            "$IsCustomProperty",
            "\"$PropertyValue\"");
    &Log ("cmd=<@cmd>", "I");
    $stat = &libconvdpf2::mysystem_array (\@cmd, 
                                     ("-timeout", "$Dpf4cCommonTimeout"));
    &Log ("status cmd=<$stat>", "I");
    if ($stat ne 0)
        {
        if ($CheckErrorText ne "")
            { $ErrorText .= $CheckErrorText . " MAINTEC informations:"; }
        $ErrorText .= &GetMaintecsap2wordErrorText ($stat);
        $ErrorText .= " Status=<$stat> Command = <@cmd>.";
        &libconvdpf2::SetDpfStatus (\%Par, 7, $ErrorText);
        exit 7;
        }
    }

# update on paper    
if ($VaiPropertyUpdateOfficeWrite eq "Y")
    {
    &Log ("Doing update on paper becaue of dpf parameter".
          "VAI_PROPERTY_UPDATE_OFFICE_WRITE=<$VaiPropertyUpdateOfficeWrite>.",
          "I");
    my @cmd = ("$PropertyUpdateExe",
            "m", # means to use runmacro
            "\"$Currfilename\"",
            "DLLSAP2Document");
    &Log ("cmd=<@cmd>", "I");
    $stat = &libconvdpf2::mysystem_array (\@cmd, 
                                     ("-timeout", "$Dpf4cCommonTimeout"));
    &Log ("status cmd=<$stat>", "I");
    if ($stat ne 0)
        {
        $ErrorText = &GetMaintecsap2wordErrorText ($stat);
        $ErrorText .= " Status=<$stat> Command = <@cmd>.";
        &libconvdpf2::SetDpfStatus (\%Par, 9, $ErrorText);
        exit 9;
        }
    }
else
    {
    &Log ("No update on paper because of working unit parameter ".
          "VAI_PROPERTY_UPDATE_OFFICE_WRITE=<$VaiPropertyUpdateOfficeWrite>",
          "I");
    }

# -------------------------------------------------------
# save all xml file names to properties
if ($Dpf4cPropertyUpdateProvideFiles eq "Y")
    { $Par{"DPF4C_PROPERTYUPDATE_OUTFILE1"} = $Par{"CURRFILENAME"}; }
# if office to office conversion is used then prevent working in
# createdestformat and modify process.
my $IsOfficeFormat = &IsOfficeFormat ($Dpf4cCommonDestFomat);
&Log ("IsOfficeFormat = <$IsOfficeFormat>.", "I");
if ($IsOfficeFormat eq "Y")
    {
    &Log ("Disable create destformats and modify process.", "I");
    $Par {"DPF4C_SWITCH_CHECKIN_ALL_DESTFORMATS"} = "N";
    $Par{"DPF4C_SWITCH_CREATE_DESTFORMATS"} = "N";
    $Par{"DPF4C_SWITCH_MODIFY"} = "N";
    $Par{"CURRFILETYPE_SAVE"} = $Currfiletype;
    $Par{"CURRFILETYPE"} = "DO_NOTHING_IN_APPKONV";
    }

# Disable property update if vai_property_update_office property
# update has been done.
my $PropUpdateVaiDone = &IsOfficeFormat ($Currfiletype);
&Log ("PropUpdateVaiDone = <$PropUpdateVaiDone>.", "I");
if ($PropUpdateVaiDone eq "Y")
    {
    &Log ("Disable create destformats and modify process.", "I");
    $Par{"DPF4C_PROPERTYUPDATE_ENABLE"} = "N";
    }

# and exit
exit $stat;

# -------------------------------------------------------
# CheckValidFormat
# -------------------------------------------------------
sub CheckValidFormat
    {
    my ($Currfiletype) = @_;
    $Currfiletype = lc ("$Currfiletype");
    # only following formats can be property updated by maintec tool
    if (!(
             ($Currfiletype eq "doc") or
             ($Currfiletype eq "docx") or
             ($Currfiletype eq "dot") or
             ($Currfiletype eq "vsd") or
             ($Currfiletype eq "xls") or
             ($Currfiletype eq "xlsx") or
             ($Currfiletype eq "xlt")))
        {
        &Log ("Info: maintec property update tool does ".
              "not support format <$Currfiletype>.", "I");
        exit 0; # Has to be 0, so other conversions (eg dwg) can run.
        }
    }

# -------------------------------------------------------
# IsOfficeFormat 
# -------------------------------------------------------
sub IsOfficeFormat  
    {
    my ($Filetype) = @_;
    my $IsOfficeFormat = "N";
    $Filetype = lc ("$Filetype");
    if (($Filetype eq "doc") or
        ($Filetype eq "docx") or
        ($Filetype eq "dot") or
        ($Filetype eq "xls") or
        ($Filetype eq "xlsx") or
        ($Filetype eq "ppt") or
        ($Filetype eq "pptx") or
        ($Filetype eq "vsd") or
        ($Filetype eq "mpv") or
        ($Filetype eq "mpp") or
        ($Filetype eq "mppx") or
        ($Filetype eq "office"))
        {
        &Log ("Filetype <$Filetype> is a office format.", "I");
        $IsOfficeFormat = "Y";
        }
    return $IsOfficeFormat;
    }

# -------------------------------------------------------
# DoOfficeRegistry - unregister and register mandatory dll files
# -------------------------------------------------------
sub DoOfficeRegistry
    {
    # under win2000 -> C:\\WINNT\\system32\\Regsvr32.exe -> regsvr32.exe
    my $RegsvrTool = $Par{"REGSVRTOOL"};
    &Log ("using RegsrvTool=<$RegsvrTool>", "I");
    # collect all dll files
    my ($Dll, @Dlls);
    push (@Dlls, $ENV{"PLSTOOLS"} . "/vai_tools/OfficePropertyUpdate/" .
          "dsofile.dll");
    push (@Dlls, $ENV{"PLSTOOLS"} . "/vai_tools/OfficePropertyUpdate/" .
          "MAINTECSAP2Word.dll");

    # check if files are available
    foreach $Dll (@Dlls)
        {
        if (!(-e $Dll))
            {
            $ErrorText = "Error: Dll file <$Dll> is not available.";
            &libconvdpf2::SetDpfStatus (\%Par, 7, $ErrorText);
            exit 11;
            }
        }

    # unregister
    foreach $Dll (@Dlls)
        {
        my $cmd_reg = "$RegsvrTool /s /u \"$Dll\"";
        my $stat = system ($cmd_reg);
        &Log ("cmd_reg=<$cmd_reg>", "I");
        if ($stat ne 0)
            {
            $ErrorText = "Error: Unregistering dll <$Dll>.";
            &libconvdpf2::SetDpfStatus (\%Par, 7, $ErrorText);
            exit 12;
            }
        }

    # register
    foreach $Dll (@Dlls)
        {
        my $cmd_reg = "$RegsvrTool /s \"$Dll\"";
        my $stat = system ($cmd_reg);
        &Log ("cmd_reg=<$cmd_reg>", "I");
        if ($stat ne 0)
            {
            $ErrorText = "Error: Registering dll <$Dll>.";
            &libconvdpf2::SetDpfStatus (\%Par, 7, $ErrorText);
            exit 13;
            }
        }
    }

# -------------------------------------------------------
# GetPropertiesAndMapVai
# -------------------------------------------------------
sub GetPropertiesAndMapVai
    {
    my ($PropertUpdateFile) = @_;
    my %HashWordProperties;

    # read properties from file
    my @PropertiesFile;
    open (PROPERTIESFILE, "$PropertUpdateFile");
    @PropertiesFile = (<PROPERTIESFILE>);
    chomp @PropertiesFile;
    close (PROPERTIESFILE);
    &Log ("Read properties from file:<@PropertiesFile>");

    my $PropertyFile;
    foreach $PropertyFile (@PropertiesFile)
        {
        my ($Key, $Value) = split /=/, $PropertyFile,2;
        if ($Key =~ /\[PROPERTIES\]/)
            { next; }
        #my $WordTitle = &GetVaiWordTitle ($Key);
        $HashWordProperties{$Key} = $Value;
        }

    return %HashWordProperties;
    }

# -------------------------------------------------------
# GetVaiWordTitle
# -------------------------------------------------------
sub GetVaiWordTitle
    {
    my ($PropertyTitle) = @_;
    my $WordTitle = $TitleHash{$PropertyTitle}; 
    if ($WordTitle eq "")
        {
        &Log ("No corresponding vai specified word title found for".
              "<$PropertyTitle>.", "W");
        $WordTitle = $PropertyTitle;      
        }
    else
        {
        &Log ("Corresponding vai specified word title for ".
              "<$PropertyTitle> is <$WordTitle>.", "I");
        }
    return $WordTitle;
    }

# -------------------------------------------------------
# GetMaintecsap2wordErrorText
# -------------------------------------------------------
sub GetMaintecsap2wordErrorText
    {
    my ($stat) = @_;
    my $ErrorText = "Error executing maintecsap2word.exe.";
    if ($stat eq 9901)
        { $ErrorText .= "Wrong method code (must be \"w\" or \"m\")."; }
    elsif ($stat eq 9902)
        { $ErrorText .= "No filename specified"; }
    elsif ($stat eq 9903)
        { $ErrorText .= "No property or macro name specified."; }
    elsif ($stat eq 9904)
        {
        $ErrorText .= "Parameter iscustomproperty is invalid ".
                     "(has to be \"y\" or \"n\")";
        }
    elsif ($stat eq 9905)
        { $ErrorText .= "No property value specified."; }
    elsif ($stat eq 9906)
        { $ErrorText .= "Wrong no of parameters."; }
    return $ErrorText;
    }

# -------------------------------------------------------
# CheckPropertyKey
# -------------------------------------------------------
sub CheckPropertyKey 
    {
    my ($UpdateProperty, $PropertyValue) = @_;
    my $CheckPropertyKey = "";
    &Log ("Checking PropertyKey <$UpdateProperty> ".
          "with PropertyValue <$PropertyValue>.", "I");
    # all custom property keys have to begin with vai
    if ($UpdateProperty =~ /^vai/i)
        {
        # no operation
        }
    else
        {
        # Check property key against some default keys
        my $SummaryOption;
        my $Exists = "N";
        foreach $SummaryOption (@SummaryOptions)
            {
            if ($SummaryOption eq $UpdateProperty)
                {
                $Exists = "Y";
                last;
                }
            }
        if ($Exists eq "N")
            {
            $CheckPropertyKey = "SEAL: Is not a valid property key, ".
                                "because the key is not a ".
                                "standard property nor a vai-property.";
            &Log ($CheckPropertyKey, "W");
            }
        }
    return $CheckPropertyKey;
    }

1;

__END__

# ---------------------------------------------------------------

=head2 NAME

 Project: DPF4Convert
 $Id: VAI_property_update_office.pl,v 1.7 2006/09/12 09:42:12 ludwig Exp $

=head2 DESCRIPTION

 Customer working unit for VAI for property update on an office document.
 Standard and customer file properties are updated and then updated 
 in the document itself (configurable if to do).
 Customer properties start with "vai". Only available properties are 
 updated. Not available properties can NOT be created.
 This working unit is supposed to be used within DPF4Convert processes.

=head2 REQUIREMENTS

=over

=item Modules

=over

=item *
modules/dpf/VAI

=back

=item Programs

=over

=item *
Executable for update and runmacro
\tools\vai_tools\OfficePropertyUpdate\MAINTECSAP2Word.exe from Maintec

=item *
Two dll files from Maintec: 
\tools\vai_tools\OfficePropertyUpdate\dsofile.dll
\tools\vai_tools\OfficePropertyUpdate\MAINTECSAP2Word.dll

=item *
installed GlobalBAO.dot - the macros are kept in this file

=back

=item Config

=over

=item *
wu-VAI_property_update_office.xml

=back

=item Documentation

=back

=head2 INPUT / OUTPUT

=over

=head3 INPUT PROPERTIES

   JOBDIR: (dir)                             Job directory.
   CURRFILENAME: (file)                      File to property udpate.
   CURRFILETYPE: (xxx)                       Current file type.
   DPF4C_PROPERTYUPDATE_ENABLE: (Y|N)        Enable property update.
   DPF4C_PROPERTYUPDATE_FILE: (file)         File with update information.
   DPF4C_PROPERTYUPDATE_PROVIDE_FILES: (Y|N) Send back updated files
   VAI_PROPERTY_UPDATE_OFFICE_WRITE: (!Y!|N) After update of propertiers even
                                             update these properties in 
                                             document too. So printing this
                                             file shows the updated properties.
   DPF4C_COMMON_DEST_FORMAT: (xxx)           Destformat
   DPF4C_COMMON_TIMEOUT: (number)            Timeout time for conversion 
                                             in minutes.
   
=head3 OUTPUT PROPERTIES

   DPF4C_PROPERTYUPDATE_OUTFILE1: (file)  Property updated office file.
   ERR_REASON: (text)              Additional error text.

=head3 PARAMETERS

 -parfile   parameter file from the DPF. Your script should
            be configured as follows:

       COMMAND="%DPFSRV%/scripts/vai/VAI_property_update_office.pl -parfile %PARFILENAME%"

=back

=head2 RETURN

 0          if OK
 != 0       if error occured
 1          error missing currfilename
 2          error missing dpf4c property update file
 3          error missing updating executable from MAINTEC 
 4          error security save of currfile
 5          error copy office file to maintec directory
 6          error move error office file to JobDir
 7          error doing property update from maintec executable
 8          error move error office file to JobDir
 9          error doing property update on paper from maintec executable
 10         error move updated office file maintec directory back
                  to job directory
 11         error dll-file missing   
 12         error unregister dll
 13         error register dll

=head2 COPYRIGHT

 Copyright 2005 SEAL Systems http://www.sealsystems.de

=cut

