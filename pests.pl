#!/usr/bin/perl
########################################################
#
# AUTHOR = Michael Vincent
# www.VinsWorld.com
#
########################################################

use vars qw($VERSION);

$VERSION = "2.3 - 27 JUL 2015";

use strict;
use warnings;
use Getopt::Long qw(:config no_ignore_case);
use Pod::Usage;

########################################################
# Start Additional USE
########################################################
use Cwd;
use Cwd 'abs_path';
use Win32::OLE;
use Win32::OLE::Const 'Microsoft Excel';
use Win32::OLE::Variant;
########################################################
# End Additional USE
########################################################

# GLOBAL variables
my ( @files, @inCells, @outCells );
my ( $outWkbk, $outWks );
my $MAXSHEETS = 50;
my $MAXROWS   = 65535;
my $MAXCOLS   = 255;

my %opt;
my ( $opt_help, $opt_man, $opt_versions );

GetOptions(
    'c|cells|incells=s'  => \$opt{inCells},
    'C|Cells|outcells=s' => \$opt{outCells},
    'debug!'             => \$opt{debug},
    'Row|outrow!'        => \$opt{outRow},
    's|sheet|insheet=i'  => \$opt{inSheet},
    'S|Sheet|outsheet!'  => \$opt{outSheet},
    'type=s'             => \$opt{inType},
    'Workbook=s'         => \$opt{outWkbk},
    'help!'              => \$opt_help,
    'man!'               => \$opt_man,
    'versions!'          => \$opt_versions,
) or pod2usage( -verbose => 0 );

pod2usage( -verbose => 1 ) if defined $opt_help;
pod2usage( -verbose => 2 ) if defined $opt_man;

if ( defined $opt_versions ) {
    print
      "\nModules, Perl, OS, Program info:\n",
      "  $0\n",
      "  Version               $VERSION\n",
      "    strict              $strict::VERSION\n",
      "    warnings            $warnings::VERSION\n",
      "    Getopt::Long        $Getopt::Long::VERSION\n",
      "    Pod::Usage          $Pod::Usage::VERSION\n",
########################################################
# Start Additional USE
########################################################
      "    Cwd                 $Cwd::VERSION\n",
      "    Win32::OLE          $Win32::OLE::VERSION\n",
########################################################
# End Additional USE
########################################################
      "    Perl version        $]\n",
      "    Perl executable     $^X\n",
      "    OS                  $^O\n",
      "\n\n";
    exit;
}

########################################################
# Start Program
########################################################

$opt{debug} = $opt{debug} || 0;

# Need a unique identifier for INSHEET to determine weather user specified it
# or we set it, as we'll do in this step
my $inSheet = $opt{inSheet} || 1;    # Default sheet is the first one

# Some output options are only allowed with output workbook specified
if ( !( defined $opt{outWkbk} ) ) {
    if ( defined $opt{outRow} ) {
        print "$0: -R requires -W\n";
        exit 1;
    }
#    if (defined $opt_outSheet) {
#        print "$0: -S requires -W\n";
#        exit 1
#    }
}

# Determine if the user spec'd an input type and validate it,
# create a variable that contains the regex search string for
# valid file types.  If nothing spec'd, then just .xls.
# Input is in form:  txt,xls ... comma separated.
if ( $opt{inType} ) {
    my @types = split /,/, $opt{inType};

    $opt{inType} = 0;

    # For each type split above into the array, match against valid types and create the REGEX.
    foreach (@types) {
        if ( !( $_ =~ /^xls(?:x)*$|^txt$|^csv$/i ) ) {
            print "$0: invalid input type -- $_\n";
        } else {
            if ( $opt{inType} ) {
                $opt{inType} = $opt{inType} . "|\\." . $_ . "\$";
            } else {
                $opt{inType} = "\\." . $_ . "\$";
            }
        }
    }
}

# If we didn't get anything or nothing provided - default to XLS
if ( !$opt{inType} ) {
    $opt{inType} = "\\.xls\$";
}

# Input file option.  Not provided at all?  Assume current directory
if ( !@ARGV ) {
    push( @ARGV, cwd . "/" );
}

# Parse input file options
foreach my $item (@ARGV) {

    # replace \ with / for compatibility with UNIX/Windows
    $item =~ s/\\/\//g;

    # Does it end with a / ?  If so, it's a directory so read files
    if ( $item =~ /\/$/ ) {

        # Since OLE is stupid and needs the full filename, we need to get the full
        # filename before we can put it in the @files array.  However, the abs_path
        # function will fail if the file doesn't exist, so we need to check if the
        # file the user provided exists first
        if ( !-e "$item" ) {
            print "$0: input directory not found - $item\n";
            exit 1;
        }

        # Get the full (c:\full\path\to\file) path
        # We need to do this now because getting the full path does
        # not provide a trialing /, so we wouldn't know if the user
        # provided a trailing slash or not.  Thus, we determine if they
        # did and then get the absolute path and then put a trailing
        # slash back on before adding to the file array.
        $item = abs_path($item);

        # Read in files from directory based on type found above
        if ( opendir( DIR, $item ) ) {
            my @ls = grep( /$opt{inType}/i, readdir(DIR) );
            closedir(DIR);

            # Add a trailing / to each item and put it in the file list
            foreach (@ls) {
                push( @files, ( $item . "/" . $_ ) );
            }
        } else {
            print "$0: input directory not found - $item\n";
            exit 1;
        }

        # Make sure we have at least 1 file to operate on
        if ( !defined $files[0] ) {
            print "$0: no files found in input directory - $item\n";
            exit;
        }

        # No trailing / , must be a file
    } else {

        # make sure individual file is acceptable type
        if ( $item =~ /\.xls(?:x)*$|\.txt$|\.csv$/i ) {

            # Since OLE is stupid and needs the full filename, we need to get the full
            # filename before we can put it in the @files array.  However, the abs_path
            # function will fail if the file doesn't exist, so we need to check if the
            # file the user provided exists first
            if ( !-e "$item" ) {
                print "$0: input file not found - $item\n";
                exit 1;
            }

            # As explained above, we need this in here due to trailing / stripping
            $item = abs_path($item);

            # replace \ with / for compatibility with UNIX/Windows
            # We'll need to do this again as the abs_path on Windows will use \
            # instead of /.
            $item =~ s/\\/\//g;

            # Does it exist - assign it as element of the @files array.
            if ( -e "$item" ) {
                push( @files, $item )

            } else {
                print "$0: input file not found - $item\n";
                exit 1;
            }

        } else {
            print "$0: input file not valid file type - $item\n";
            exit 1;
        }
    }
}
# DONE! Parse input file options

# Parse input Cell option
if ( defined $opt{inCells} ) {

    # It must be a number range in the form:
    #   Z!x,X:y-Y; (etc...)
    # (optional NUMBER!)NUMBER(optional '-' or ',' NUMBER):NUMBER(optional '-' or ',' NUMBER)(optional ;[repeat])
    if ($opt{inCells} !~ /^((\d+\!)*\d+([\,\-]\d+)*:\d+([\,\-]\d+)*(\;)?)+$/ )
    {
        print "$0: in range only allows number, '!', ':', '-' or ','\n";
        exit 1;
    }

    # Format is OK, get the cells required by parsing the argument
    # and get the output in the @inCells array
    @inCells = &GET_RANGE_ARGS( $inSheet, $opt{inCells} );
}

# Parse output Cell option
if ( defined $opt{outCells} ) {

    # This option requires -W out workbook.  Otherwise, we have no control over where we output things.
    if ( !( defined $opt{outWkbk} ) ) {
        print "$0: -C requires -W\n";
        exit 1;
    }

    # It must be a number range in the form:
    #   Z!x,X:y-Y; (etc...)
    # (optional NUMBER!)NUMBER(optional '-' or ',' NUMBER):NUMBER(optional '-' or ',' NUMBER)(optional ;[repeat])
    #
    # NOTE:  We used to not allow (optional NUMBER!)
    #    if ($opt_outCells !~ /^(\d+([\,\-]\d+)*:\d+([\,\-]\d+)*(\;)?)+$/) {
    #        print "$0: out range only allows number, ':', '-' or ','\n";
    #
    if ( $opt{outCells}
        !~ /^((\d+\!)*\d+([\,\-]\d+)*:\d+([\,\-]\d+)*(\;)?)+$/ ) {
        print "$0: out range only allows number, '!', ':', '-' or ','\n";
        exit 1;
    }

    # Format is OK, get the cells required by parsing the argument
    # and get the output in the @outCells array.
    # Use '1' as the default sheet if none specified in the -C argument pieces
    @outCells = &GET_RANGE_ARGS( 1, $opt{outCells} );

    # Verify 1 to 1 mapping of in/out cells.  We can't have more in or out cells or we
    # won't know where to put things.  The exception is if there is only 1 outCell provided.
    # In this case, we can use it as a starting point for output and build on that.
    if ( @outCells != @inCells ) {
        if ( @outCells != 1 ) {
            print "$0: -c in cell count != -C out cell count\n";
            exit 1;
        }
    }
}

# Create the Excel object
my $Excel = Win32::OLE->GetActiveObject('Excel.Application')
  || Win32::OLE->new( 'Excel.Application', 'Quit' );
$Excel->{DisplayAlerts} = 0;

# Parse output file name
# MUST do this after input is parsed so that the output file is not
# created and then becomes part of a possible directory listing for
# the input file parsing.  If we're debugging, we just need the filename,
# we don't actually do anything to it, so we can skip this part.
if ( defined $opt{outWkbk} ) {

    # Does it start with a \ ?  If so, add the drive otherwise, it must
    # be off the local directory, so add the current path
    if ( !( $opt{outWkbk} =~ /^[A-Za-z]:/ ) ) {
        if ( $opt{outWkbk} =~ /^\\/ ) {

            my @drive = split /:/, cwd;
            $opt{outWkbk} = $drive[0] . ":" . $opt{outWkbk};
        } else {
            $opt{outWkbk} = cwd . "\\" . $opt{outWkbk};
        }
    }

    # Replace all / with \ (must be \\ to escape the \) to make Windows compatible
    $opt{outWkbk} =~ s/\//\\/g;

    # The following commands are stuff tried to get the broken Win32::OLE SaveAs method to work:
    #
    # $opt{outWkbk} =~ s/\\/\\\\/g;
    # $opt{outWkbk} = "'" . $opt{outWkbk} . "'";

    # Done all our massaging.  Now it's time to open the file.  Check if it exists or if we
    # need to create it.
    # Does it exist?  If so, open it ...
    if ( -e "$opt{outWkbk}" ) {

        if ( !$opt{debug} ) {
            print "WARN:  Output file exists - $opt{outWkbk}\n";
            $outWkbk = $Excel->Workbooks->Open( $opt{outWkbk} );
        }

        # Otherwise, create a new book.
    } else {

        if ( !$opt{debug} ) {

            # At this point, $opt{outWkbk} is in the form:
            #  DRIVE:\path\to\file.ext
            #
            # Normally, we'd be done here, but MS Excel SaveAs feature is STUPID
            # and needs a fully qualified path name.  But of course, when you do that
            # you have something like C:\path\to\file.name and Excel complains about a
            # : in the filename.  Thus, by default it saves in "my documents", so we can
            # ..\..\..\ escape to root and then give our fully qualifed path minus
            # the drive letter and :.  This of couse will break any mapped drive access.
            # I'm working on this ...
            #
            # And now, 2 days later, it's magically fixed?  Keep this here ... just in case ;-)
#            if ($opt{outWkbk} =~ /:/) {
#                my @temp = split /:/, $opt{outWkbk};
#                $opt{outWkbk} = "..\\..\\.." . $temp[1]
#            }

            # Create new output file (as per normal)
            $outWkbk = $Excel->Workbooks->Add();
            $outWkbk->SaveAs(
                {   Filename   => $opt{outWkbk},
                    FileFormat => xlWorkbookNormal
                }
            );
        }
    }

    if ( !$opt{debug} ) {

        # Assign the active worksheet.  We'll start with '1' so there'll be a value in this
        # outWks variable, but it will change later as needed
        $outWks = $outWkbk->Worksheets(1);
    }
}

########################################################
#################################
# DONE args, do STUFF           #
#################################

my $outSheet = 1;    # Initial value of output Sheet
my $outRow   = 1;    # Initial value of output Row
my $outCol   = 1;    # Initial value of output Col
my $sheetCount = 0; # Keeps track of which sheet we're operating on in output to increase as needed
my $fileLoopCount = 0; # Keeps track of which file we're operating on to increase output by row or col

# Loop through each file
foreach my $file (@files) {

    # We'll check to see if the input file is the same as the output file.
    # This can occur if you run it once in a directory and create an output file
    # and then run the same command in the same directory.  The output file you
    # are specifying will now exist and it will also be used as input and there
    # will be a "recursion" issue.  If that happens, we'll just skip that file.
    #
    # First off is to convert / back to \ in the infile name, then compare ...
    $file =~ s/\//\\/g;
    if ( defined $opt{outWkbk} ) {

        if ( $file eq $opt{outWkbk} ) {
            if ( !$opt{debug} ) {
                print
                  "WARN:  input file $file is same as output file $opt{outWkbk} ... Skipping ...\n";
            }
            next;
        }
        if ( !$opt{debug} ) {
            print "INFO:  processing input file - $file\n";
        }
    }

    # Open the input file
    my $inWkbk = $Excel->Workbooks->Open($file);

    # Print file name if we're debugging
    if ( $opt{debug} ) {
        print "FILE = $file\n";
    }

    # reset counters for next iteration
    if ( defined $opt{outRow} ) {
        $outRow = 1;
    } else {
        $outCol = 1;
    }
    my $outCellArrayCount = 0;

    if ( ( !( defined $opt{outWkbk} ) ) and ( defined $opt{outSheet} ) ) {
        print "FILE = $file\n";
        for ( 1 .. $inWkbk->Worksheets->Count() ) {
            my $inWks = $inWkbk->Worksheets($_);
            printf "  Sheet $_ = %s\n", $inWks->{Name};
        }

        # If inCells were provided, loop through the requested cells and pull the values to print ...
    } elsif (@inCells) {

        foreach my $inCell (@inCells) {

            # Arrange inCell coordinates
            # Split each coordinate into sheet/coord and each coord into row then column
            my (@inCoords) = split /!/, $inCell;
            ( $inCoords[1], $inCoords[2] ) = split /:/, $inCoords[1];

            my $inWks = 0;
            my $inWkc = 0;

            # Assign the working cell
            if ( $inWkbk->Worksheets->Count() < $inCoords[0] ) {
                $inWkc = "<UNDEF>";
            } else {

                # MS Excel is STUPID, in that a number is not a number unless you tell it
                # via some stupid VARIANT thing.
                $inWks
                  = $inWkbk->Worksheets( Variant( VT_I4, $inCoords[0] ) );
                $inWkc = $inWks->Cells(
                    Variant( VT_I4, $inCoords[1] ),
                    Variant( VT_I4, $inCoords[2] )
                )->{Value};
            }

            # Arrange out cell coordinates (if provided)
            if (@outCells) {

                # In this case, only 1 outCell is provided and we'll anchor output starting there.
                # Thus, we'll always use $outCells[0] - the first (and only) element in the outCells
                # array as our operator.
                if ( @outCells == 1 ) {

                    # Arrange outCell coordinates
                    # Split each coordinate into sheet/coord and each coord into row then column
                    my (@outCoords) = split /!/, $outCells[0];
                    ( $outCoords[1], $outCoords[2] ) = split /:/,
                      $outCoords[1];

                    $outSheet = $outCoords[0];

                    # Have we added sheets that go beyond the existing in outbook?
                    # If so, we need to create the new sheet before we print to it
                    if ( !$opt{debug} ) {
                        if ( $outSheet > $outWkbk->Worksheets->Count() ) {
                            &AddSheets( $outSheet, $opt{outWkbk}, $outWkbk );
                        }
                        $outWks = $outWkbk->Worksheets(
                            Variant( VT_I4, $outSheet ) );
                    }

                    # Assign the working cell
                    #
                    # We need to remember we're offset by the outCells[0] value.  So we'll need to
                    # add that to our outRow/Col.  Also, we'll need to keep track of which value in
                    # the input file we're writing ($outCellArrayCount can do this as it counts along
                    # which output cell we'd be using) and which file we're on to increase our row
                    # (or col if -R provided) (fileLoopCount can do this).
                    if ( defined $opt{outRow} ) {
                        $outCol = $fileLoopCount + $outCoords[2];
                        $outRow = $outCellArrayCount + $outCoords[1];
                    } else {
                        $outRow = $fileLoopCount + $outCoords[1];
                        $outCol = $outCellArrayCount + $outCoords[2];
                    }

                    # Otherwise, we'll need to keep track of which element in the outCells array we're
                    # operating on by the outCellArrayCount variable.
                } else {

                    # Arrange outCell coordinates
                    # Split each coordinate into sheet/coord and each coord into row then column
                    my (@outCoords) = split /!/,
                      $outCells[$outCellArrayCount];
                    ( $outCoords[1], $outCoords[2] ) = split /:/,
                      $outCoords[1];

                    $outSheet = $outCoords[0];

                    # Have we added sheets that go beyond the existing in outbook?
                    # If so, we need to create the new sheet before we print to it
                    if ( !$opt{debug} ) {
                        if ( $outSheet > $outWkbk->Worksheets->Count() ) {
                            &AddSheets( $outSheet, $opt{outWkbk}, $outWkbk );
                        }
                        $outWks = $outWkbk->Worksheets(
                            Variant( VT_I4, $outSheet ) );
                    }

                    # Assign the working cell
                    $outRow = $outCoords[1];
                    $outCol = $outCoords[2];

                    # We need to increment by column or row the outcell count per new file
                    # This is determind by default = Row or is -R, then by column
                    if ( ( $opt{outRow} ) ) {
                        $outCol = $outCol + $fileLoopCount;
                    } else {
                        $outRow = $outRow + $fileLoopCount;
                    }

                }

                # Move to next element in out array
                $outCellArrayCount++;
            }

            &PrintAll( $inCoords[0], $inCoords[1], $inCoords[2], $inWkc,
                $outSheet, $outRow, $outCol, $outWks, );

            # Do we increment by row or column?
            if ( defined $opt{outRow} ) {
                $outRow++;
            } else {
                $outCol++;
            }
        }    # foreach @incells

        # Otherwise, inCells were not provided, do the whole file, sheet at a time
    } else {

        # Default start sheet is 1 and max sheet is the highest sheet in the workbook
        my $startSheet = 1;
        my $lastSheet  = $inWkbk->Worksheets->Count();

        # If the user specified a default sheet ...
        if ( defined $opt{inSheet} ) {

            # That is greater than the number of sheets in this file, no use printing errors
            # for everything, just skip this file and go the NEXT one in the foreach files loop ...
            if ( $opt{inSheet} > $lastSheet ) {
                next

                  # Otherwise, the user specified sheet is the only one we operate on
            } else {
                $startSheet = $lastSheet = $opt{inSheet};
            }
        }

        # Previous checking guarantees that we'll only see outCells here if it is 1, but we'll
        # check again anyway rather than just saying (@outCells).  If it is provided at this point,
        # we're using it as a starting anchor point for output.  See above comments relating to this.
        if ( @outCells == 1 ) {
            my (@outCoords) = split /!/, $outCells[0];
            ( $outCoords[1], $outCoords[2] ) = split /:/, $outCoords[1];

            # If we're in sheet mapping mode, we use the outcoords as a starting anchor point
            # and we're just mapping so no increase by filecount or outcoords

            $outSheet = $sheetCount + $outCoords[0];

            # Have we added sheets that go beyond the existing in outbook?
            # If so, we need to create the new sheet before we print to it
            if ( !$opt{debug} ) {
                if ( $outSheet > $outWkbk->Worksheets->Count() ) {
                    &AddSheets( $outSheet, $opt{outWkbk}, $outWkbk );
                }
                $outWks = $outWkbk->Worksheets( Variant( VT_I4, $outSheet ) );
            }

            if ( defined $opt{outSheet} ) {
                $outRow = $outCoords[1];
                $outCol = $outCoords[2]

            } else {

                # Assign the working cell
                if ( defined $opt{outRow} ) {
                    $outCol = $fileLoopCount + $outCoords[2];
                    $outRow = $outCellArrayCount + $outCoords[1];
                } else {
                    $outRow = $fileLoopCount + $outCoords[1];
                    $outCol = $outCellArrayCount + $outCoords[2];
                }
            }
        }

        # There are no input cells if we're here, so we make our own by iterating over
        # all values in all sheets
        for ( my $i = $startSheet; $i <= $lastSheet; $i++ ) {

            # Assign the current working sheet
            my $inWks = $inWkbk->Worksheets($i);

            # We always read by rows - find min/max and iterate them

            # BUT FIRST, we need to make sure we can find a min row, otherwise there is
            # nothing in the sheet and we'll get errors.  So, if an empty sheet is found
            # skip it.
            if (!(  defined $inWks->UsedRange->Find(
                        {   What            => "*",
                            SearchDirection => xlNext,
                            SearchOrder     => xlByRows
                        }
                    )
                )
              ) {
                next;
            }
            my $MinRow = $inWks->UsedRange->Find(
                {   What            => "*",
                    SearchDirection => xlNext,
                    SearchOrder     => xlByRows
                }
            )->{Row};
            my $MaxRow = $inWks->UsedRange->Find(
                {   What            => "*",
                    SearchDirection => xlPrevious,
                    SearchOrder     => xlByRows
                }
            )->{Row};

            for ( my $iR = $MinRow; defined $MaxRow && $iR <= $MaxRow; $iR++ )
            {

                # Next, we read by cols - find min/max and iterate them
                my $MinCol = $inWks->UsedRange->Find(
                    {   What            => "*",
                        SearchDirection => xlNext,
                        SearchOrder     => xlByColumns
                    }
                )->{Column};
                my $MaxCol = $inWks->UsedRange->Find(
                    {   What            => "*",
                        SearchDirection => xlPrevious,
                        SearchOrder     => xlByColumns
                    }
                )->{Column};

                for (
                    my $iC = $MinCol;
                    defined $MaxCol && $iC <= $MaxCol;
                    $iC++
                  ) {

                    # Assign working cell
                    my $inWkc = $inWks->Cells( $iR, $iC )->{Value};

                    # If sheet mapping mode, we need to set the output anchor point as the input
                    # cell (mapping) plus any output anchor
                    if ( defined $opt{outSheet} ) {

                        # Have we added sheets that go beyond the existing in outbook?
                        # If so, we need to create the new sheet before we print to it
                        if ( !$opt{debug} ) {
                            if ( $outSheet > $outWkbk->Worksheets->Count() ) {
                                &AddSheets( $outSheet, $opt{outWkbk},
                                    $outWkbk );
                            }
                            $outWks = $outWkbk->Worksheets(
                                Variant( VT_I4, $outSheet ) );
                        }

                        $outRow = $outRow + $iR - 1;
                        $outCol = $outCol + $iC - 1;
                    }

                    &PrintAll( $i, $iR, $iC, $inWkc,
                        $outSheet, $outRow, $outCol, $outWks, );

                    # If we're in sheet mapping mode, we need to reset our anchor point
                    if ( defined $opt{outSheet} ) {
                        $outRow = $outRow - $iR + 1;
                        $outCol = $outCol - $iC + 1

                    } else {

                        # Do we increment by row or column?
                        if ( defined $opt{outRow} ) {
                            $outRow++;
                        } else {
                            $outCol++;
                        }
                    }

                }    # for columns
            }    # for rows

            # If we're in sheet mapping mode, we need to create a new sheet if the in book
            # has another sheet to do
            if ( defined $opt{outSheet} ) {
                $outSheet++;
                $sheetCount++;
            }

        }    # for sheets
    }    # else

    # We're done with the input file, so close it before going on to the next one
    $inWkbk->close();

    # Get ready for next Workbook
    # Do we increment by row or column?  By default, we're adding to the row, unless -R specified
    if ( !defined $opt{outSheet} ) {
        if ( defined $opt{outRow} ) {
            $outCol++;
        } else {
            $outRow++;
        }
    }

    # Add one to the file loop count as we're going on to the next file
    $fileLoopCount++;
    print "\n" if ( !( defined $outWkbk ) )

}
# DONE! All files parsed after loop complete

if ( defined $outWkbk ) {

    $outWkbk->SaveAs(
        {   Filename   => $opt{outWkbk},
            FileFormat => xlWorkbookNormal
        }
    );
    $outWkbk->close();
}

########################################################
#################################
# DONE                          #
#################################

########################################################
#################################
# Print All                     #
#################################
sub PrintAll () {

    use strict;

    my ( $iS, $iR, $iC, $inWkc, $outSheet, $outRow, $outCol, $outWks, ) = @_;

    # Debug printing? ...
    if ( $opt{debug} ) {
        print "$iS!$iR:$iC = ";
        print $inWkc if ( defined $inWkc );

        if ( defined $opt{outWkbk} ) {

            # It would be brilliant if this worked; however, since Win32::OLE SUCKS and needs the
            # full filename and the rest of the world can deal with relative paths, I need to
            # convert the possible relative path of the outworkbook provided and compare that to the
            # fully qualifed path in file we're looping on.  I'll try to fix this later.
            my @filename = split /\\/, $opt{outWkbk};

            printf "\t-> %s", $filename[scalar(@filename) - 1];
            print "(", $outSheet, "!", $outRow, ":", $outCol, ")";
        }
        print "\n"

          # Otherwise, normal output printing
    } else {

        # Normal printing to output workbook
        if ( defined $opt{outWkbk} ) {

            $outWks->Cells( Variant( VT_I4, $outRow ),
                Variant( VT_I4, $outCol ) )->{Value} = $inWkc
              if ( defined $inWkc )

              # Normal printing to STDOUT
        } else {

            # Print if value exists
            if ( defined $inWkc ) {
                print $inWkc;
            }
            print "\t";
        }
    }
}

########################################################
#################################
# Add Sheets                    #
#################################
sub AddSheets () {

    use strict;

    my ( $outS, $outW, $outWkbk ) = @_;

    print "WARN:  Worksheet $outS does not exist in output file - $outW\n";
    while ( $outS > $outWkbk->Worksheets->Count() ) {
        $outWks
          = $outWkbk->Worksheets->Add(
            {after => $outWkbk->Worksheets( $outWkbk->Worksheets->{count} )}
          );
        printf "INFO:  Creating worksheet %i in output file - $outW\n",
          $outWkbk->Worksheets->Count();
    }
}

########################################################
#################################
# Get RANGE Arguments           #
#################################
# This sub handles arguments that can have several values with
# dash '-' and comma ',' as operators
#
#   Z!x,X:y-Y; (etc...)
#
# (optional NUMBER!)NUMBER(optional '-' or ',' NUMBER):NUMBER(optional '-' or ',' NUMBER)(optional ;[repeat])
#
# To get all the options correctly, we need to have this option
# accept an optional string value and then parse the string
# for dash '-' and comma ','
#
# RETURN:  returns an array populated with cell data in the form:
#          Z!X:Y
#          Where Z is sheet, X is row and Y is column.  We use the ! and : to split later
sub GET_RANGE_ARGS () {

    use strict;

    my ( $sheet, $opt ) = @_;

    my ( @final, @temp, @sheet, @row, @col );

    # Split the string at the commas first to get
    my (@option) = split /;/, $opt;

    # We need a row/column counter for the main loop
    # We're going to split each option at the row:col demark
    # then operate on each part each time through the loop (twice)
    # 0 = row, 1 = col
    my $j = 0;
    foreach my $group (@option) {

        if ( $group !~ /!/ ) {
            $group = $sheet . "!" . $group;
        }

        # Split the group into sheet, row and column
        my (@coords) = split /!/, $group;
        ( $coords[1], $coords[2] ) = split /:/, $coords[1];

        # Reset counters and arrays for the new group
        my $SRC = 0;
        @sheet = ();
        @row   = ();
        @col   = ();
        foreach my $coord (@coords) {

            # Split each coordinate by commas
            my (@value) = split /,/, $coord;

            # Now we'll loop through the remaining values to see if there are
            # dashes.  Dashes means all numbers between, inclusive.  Thus, we'll
            # need to expand the ranges and put the values in the array.
            my $i = 0;
            @temp = ();
            foreach my $value (@value) {

                # If the value we're looking at has a dash '-', then we'll split
                # and add the 'missing' numbers.
                if ( $value =~ /-/ ) {

                    # what if they value is 1-4-7, I can't catch this with the regex to
                    # admit to this procedure, so we just get the length of the array
                    # after the split and assign that to a variable that we'll use as the
                    # upper range limit
                    my (@startEnd) = split /-/, $value;

                    if ( $startEnd[0] <= $startEnd[$#startEnd] ) {

                        # Iterate through by row then column and create the coordinates
                        for (
                            my $start = $startEnd[0];
                            $start <= $startEnd[$#startEnd];
                            $start++
                          ) {
                            $temp[$i++] = $start;
                        }

                    } elsif ( $startEnd[0] > $startEnd[1] ) {

                        for (
                            my $start = $startEnd[0];
                            $start >= $startEnd[$#startEnd];
                            $start--
                          ) {
                            $temp[$i++] = $start;
                        }

                        # We should never get here
                    } else {
                        print "$0: Skipping invalid range:  $opt\n";
                    }

                    # If the current $value doesn't have a dash '-', then just move on
                } else {
                    $temp[$i++] = $value;
                }
            }

            # did we populate @temp array with sheets, rows or columns?
            if    ( $SRC == 0 ) { @sheet = @temp }
            elsif ( $SRC == 1 ) { @row   = @temp }
            else                { @col   = @temp }

            # Move on to rows then columns
            $SRC++;
        }

        # Add new sheet!row:col values to final array
        foreach my $sheet (@sheet) {
            foreach my $row (@row) {
                foreach my $col (@col) {

                    if ( $sheet > $MAXSHEETS ) {
                        print "$0: Found sheet [$sheet] > max [$MAXSHEETS]\n";
                        exit;
                    } elsif ( $row > $MAXROWS ) {
                        print "$0: Found row [$row] > max [$MAXROWS]\n";
                        exit;
                    } elsif ( $col > $MAXCOLS ) {
                        print "$0: Found col [$col] > max [$MAXCOLS]\n";
                        exit;
                    } else {
                        $final[$j++] = $sheet . "!" . $row . ":" . $col;
                    }
                }
            }
        }
    }

    # return the values of the temp arrays after reconnecting row:col format
    return (@final);
}

########################################################
# End Program
########################################################

__END__

########################################################
# Start POD
########################################################

=head1 NAME

PESTS - Parse Excel Spreadsheets To Single

=head1 SYNOPSIS

 pests [options] [[files | dir/] ... ]

=head1 DESCRIPTION

Script parses Excel (.XLS[X]) files in a given directory and
extracts the provided cells into a single output, one row
for each parsed sheet with the columns equal to the values
of the cells in the provided command line argument.

=head1 OPTIONS

=head2 INPUT OPTIONS

 files       Optional Input Excel Workbook.  If this ends in a "/", 
 dir/        assumed to be a directory containing Excel files, all of 
             which will be parsed.
             DEFAULT:  (or not specified) Local directory and only XLS
                       files.  To control input file type when specifying 
                       directory, see -t option.

 -c cells    Input cells read from input workbook.  String of cells to 
 --cells     put into the master output in format:

                [SHEET!]ROW[(-|,)ROW']:COL[(-|,)COL'][; ...]

             Where SHEET is the optional sheet number in the Excel 
             wookbook.
             DEFAULT:  (or not specified) 1.

             Also, ROW is the row number in the Excel sheet and COL is 
             the column number in the Excel sheet.  Note that columns in
             Excel use letters, so in this program, A=1, B=2, C=3 ... 
             and so on.

             Use '-' for ranges and ',' separating non-contiguous values.
             Each cell group is separated by ';'.

             Ranges will iterate across columns, then down rows.  For 
             example:

                1-3:4-6

             Evalutes to:

                Row1:Col4,Row1:Col5,Row1:Col6,Row2:Col4 ...

             NOT:

                Row1:Col4,Row2:Col4,Row3:Col4,Row1:Col5 ...

             DEFAULT:  (or not specified) Script will assume all cells 
                       on all sheets for all input files and parse 
                       accordingly.

 -s #        Default worksheet of input Excel Workbook.  With -c, if 
 --sheet     SHEET! is not specified, default is 1.  This option allows 
             the changing of the default sheet.

 -t type     Input file type.  When an input directory is specified, the 
 --type      default file type is assumed to be .XLS.  To modify, use 
             this option.  Valid types are:

                XLS  = Excel spreadsheet (97 - 2003)
                XLSX = Excel spreadsheet (2007 - )
                TXT  = Tab delimited
                CSV  = Comma Separated Values

             Arguments are a comma-separated list.  For example, to
             include only CSV and XLS files in all input directory 
             arguments, use:

                -t csv,xls

=head2 OUTPUT OPTIONS

 -W wkbk     Output Excel workbook in .XLS format.
 --Workbook  DEFAULT:  (or not specified) Output tab-delimited text to 
                       STDOUT.

=head3 SUB-OPTIONS

   -C cells  Output cells to write in output workbook.  String follows 
   --Cells   same format for input cells (defined above).

             Number of output cells must match number of input cells or 
             an error is generated and program stopped.  The only 
             exception is if a single output cell is provided and none 
             or more than one input cells are provided.  In this case, 
             provided output cell is used as starting point for output.

             Start at Row1:Col1 then Row1:Col2 ...
             Increment Row for each new input file.

             If cells are specified, all input cells from first 
             input file are written to the output cells specified.
             For each next file, the input cells are stored in 
             output cells Row+1.  For example:

                INPUT     OUTPUT FILE
                File 1 -> RowA:ColB then RowX:ColY, ...
                File 2 -> RowA+1:ColB then RowX+1:ColY, ...
                File 3 -> RowA+2:ColB then RowX+2:ColY, ...

   -R        Increment output cells down rows then across columns for 
   --Row     each new input file.
             DEFAULT:  (or not specified) Increment output cells across 
                       columns then down rows for each new input file.

             When output cells are specified, the default behavior is as 
             described above.  This option changes behavior to:

                INPUT     OUTPUT FILE
                File 1 -> RowA:ColB then RowX:ColY, ...
                File 2 -> RowA:ColB+1 then RowX:ColY+1, ...
                File 3 -> RowA:ColB+2 then RowX:ColY+2, ...

   -S        Increment output cells to new worksheets for each input 
   --Sheet   file.  Use -C options to specify starting locations.

             -S without -W option simply prints the names of all sheets 
             in the input workbooks.

 -d          Debug output.  Overrides output file if provided.  Output 
 --debug     is sent as plain text to STDOUT in form:

                SHEET!ROW:COL = VALUE [-> OUT_FILE_NAME(SHEET!ROW:COL)]

             [<text>] is specified only if an output file is given 
             although the output file is not actually generated.

 --help      Print Options and Arguments.
 --man       Print complete man page.
 --versions  Print Modules, Perl, OS, Program info.

=head1 LICENSE

This software is released under the same terms as Perl itself.
If you don't know what that means visit L<http://perl.com/>.

=head1 AUTHOR

Copyright (C) Michael Vincent 2008

L<http://www.VinsWorld.com>

All rights reserved

=cut
