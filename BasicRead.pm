#
#  Spreadsheet::BasicRead.pm
#
#  Synopsis:            see POD at end of file
#
#
#-- The package
#--------------------------------------------------
package Spreadsheet::BasicRead;

$VERSION = sprintf("%d.%02d", q'$Revision: 1.1.1.1 $' =~ /(\d+)\.(\d+)/);
#--------------------------------------------------
#
#
my $CVS_Log = q{

$Log

};
#
#
#


#-- Define the locations of the ARS and related modules
use lib qw(
	.
	/opt/ar/perl/lib/perl5/5.6.0/aix
	/opt/ar/perl/lib/perl5/5.6.0
	/opt/ar/perl/lib/perl5/site_perl/5.6.0/aix
	/opt/ar/perl/lib/perl5/site_perl/5.6.0
	/opt/ar/perl/lib/perl5/site_perl
);


#-- Required Modules
#-------------------
use strict;
use warnings;
use Spreadsheet::ParseExcel;



#-- Linage
#---------
our @ISA = qw( Spreadsheet::ParseExcel );


sub new
{
	my $proto  = shift;
	my $class  = ref($proto) || $proto;

	my $self = {};
	bless($self, $class);

	$self->{skipBlankRows} = 0;

	# Do we have any arguments to process
	#------------------------------------

	# Is there just one argument?  If so treat as filename, otherwise assume named arguments
	if (@_ == 1)
	{
		$self->{fileName} = $_[0];
		$self->openSpreadsheet($self->{fileName});

		return $self;
	}



	# If we get to here then we assume named arguments to process
	my %args = @_;

	# Is there a log object
	if (defined($args{log}) && $args{log} ne '')
	{
		$self->{log} = $args{log};
	}

	# Do we skip blank rows
	if (defined($args{skipBlankRows}))
	{
		$self->{skipBlankRows} = $args{skipBlankRows} ? 1 : 0;
	}

	# Is there a file to open
	if (defined($args{fileName}) && $args{fileName} ne '')
	{
		$self->{fileName} = $args{fileName};
		$self->openSpreadsheet($args{fileName});
	}

	return $self;
}



sub openSpreadsheet
{
	my ($self, $ssFileName) = @_;

	#-- Open the Excel spreadsheet and process
	my $ssExcel = new Spreadsheet::ParseExcel;
	my $ssBook  = $ssExcel->Parse($ssFileName);
	unless ($ssBook)
	{
		$self->logexp("Could not open Excel spreadsheet file '$ssFileName': $!");
	}

	# Store the objects
	$self->{ssExcel} = $ssExcel;
	$self->{ssBook}  = $ssBook;

	# Get the first sheet
	$self->getFirstSheet();

	return ($ssExcel, $ssBook)
}




sub numSheets
{
	my $self = shift;

	return defined($self->{ssBook}) ? $self->{ssBook}->{SheetCount} : undef;
}



sub currentSheetNum
{
	my $self = shift;

	return defined($self->{currentSheetNum}) ? $self->{currentSheetNum} : undef;
}



sub currentSheetName
{
	my $self = shift;

	return defined($self->{ssSheet}) ? $self->{ssSheet}->{Name} : undef;
}



sub setCurrentSheetNum
{
	my $self = shift;

	return $self->{currentSheetNum} = $_[0];
}



sub getNextSheet
{
	my $self = shift;

	my $currentSheet = $self->currentSheetNum();

	# No sheet, so get the first sheet
	return $self->getFirstSheet() unless (defined($self->{ssSheet}));

	# Get the next sheet
	if (defined($self->{ssSheet}) && $currentSheet < $self->numSheets())
	{
		$self->setCurrentSheetNum = ++$currentSheet;
		$self->{ssSheet}    = $self->{ssBook}->{Worksheet}[$currentSheet];
		$self->{ssSheetRow} = $self->{ssSheet}->{MinRow} if (defined($self->{ssSheet}));
		$self->{ssSheetCol} = $self->{ssSheet}->{MinCol} if (defined($self->{ssSheet}));
		return $self->{ssSheet};
	}

	return undef;
}



sub getFirstSheet
{
	my $self = shift;

	$self->{setCurrentSheetNum} = 0;
	$self->{ssSheet}    = $self->{ssBook}->{Worksheet}[0] if (defined($self->{ssBook}));
	$self->{ssSheetRow} = -7;  # Flag to getNextRow that this is the first row
	$self->{ssSheetCol} = $self->{ssSheet}->{MinCol}      if (defined($self->{ssSheet}));
	return $self->{ssSheet};
}


sub cellValue
{
	my ($self, $r, $c) = @_;
	return undef unless (defined($self->{ssSheet}) && defined($self->{ssSheet}->{Cells}[$r][$c]));
	return $self->{ssSheet}->{Cells}[$r][$c]->Value;
}



sub getFirstRow
{
	my $self = shift;

	return undef unless defined($self->{ssSheet});

	my $row = $self->{ssSheet}->{MinRow};
	$self->{ssSheetRow} = $row;


	# Loop through each column and put into array
	my $x     = 0;
	my @data  = ();
	my $blank = 0;
	for (my $col = $self->{ssSheet}->{MinCol}; $col <= $self->{ssSheet}->{MaxCol}; $x++, $col++)
	{
		no warnings qw(uninitialized);

		# Note that this is the formatted value of the cell (ie what you see, no the real value)
		$data[$x] = $self->cellValue($row, $col);

		# remove leading and trailing whitespace
		$data[$x] =~ s/^\s+//;
		$data[$x] =~ s/\s+$//;
		$blank++ unless $data[$x] =~ /^$/;
	}


	return ($self->{skipBlankRows} && $blank == 0) ? $self->getNextRow() : \@data;
}



sub getNextRow
{
	my $self = shift;

	# Must have a sheet defined
	return undef unless defined($self->{ssSheet});

	# Find the next row and make sure it's valid
	my $row = ++$self->{ssSheetRow};
	return undef if ($row > $self->{ssSheet}->{MaxRow});

	# If row is zero or negative then this is the first row
	return $self->getFirstRow() if ($row <= 0);


	# Loop through each column and put into array
	my $x     = 0;
	my @data  = ();
	my $blank = 0;
	for (my $col = $self->{ssSheet}->{MinCol}; $col <= $self->{ssSheet}->{MaxCol}; $x++, $col++)
	{
		no warnings qw(uninitialized);

		# Note that this is the formatted value of the cell (ie what you see, no the real value)
		$data[$x] = $self->cellValue($row, $col);

		# remove leading and trailing whitespace
		$data[$x] =~ s/^\s+//;
		$data[$x] =~ s/\s+$//;
		$blank++ unless $data[$x] =~ /^$/;
	}

	return ($self->{skipBlankRows} && $blank == 0) ? $self->getNextRow() : \@data;
}



sub logexp
{
	my $self = shift;

	my $msg = join('', @_);
	if (defined $self->{log})
	{
		$self->{log}->exp($msg);
	}

	die $msg;
}



sub logmsg
{
	my $self  = shift;
	my $level = shift;

	my $msg = join('', @_);
	if (defined $self->{log})
	{
		$self->{log}->msg($level, $msg);
	}
	else
	{
		print STDERR $msg;
	}
}



#####################################################################
# DO NOT REMOVE THE FOLLOWING LINE, IT IS NEEDED TO LOAD THIS LIBRARY
1;

__END__

## POD DOCUMENTATION ##


=head1 NAME

Spreadsheet::BasicRead - Methods to easily read data from spreadsheets


=head1 DESCRIPTION

Provides methods for simple reading of a Excel spreadsheet row
at a time returning the row as an array of column values.
Properties can be set so that blank rows are skipped


=head1 SYNOPSIS

 use Spreadsheet::BasicRead;

 my $xlsFileName = 'Test.xls';

 my $ss = new Spreadsheet::BasicRead($xlsFileName) ||
 	die "Could not open '$xlsFileName': $!";

 # Print the row number and data for each row of the
 # spreadsheet to stdout using '|' as a separator
 my $row = 0;
 while (my $data = $ss->getNextRow())
 {
 	$row++;
 	print join('|', $row, @$data), "\n";
 }

 # Print the number of sheets
 print "There are ", $ss->numSheets(), " in the spreadsheet\n";

 # Set the heading row to 4
 $ss->setHeadingRow(4);

 # Skip the first data line, it's assumed to be a heading
 $ss->skipHeadings(1);

 # Print the name of the current sheet
 print "Sheet name is ", $ss->currentSheetName(), "\n";

 # Reset back to the first row of the sheet
 $ss->getFirstRow();


=head1 REQUIRED MODULES

The following modules are required:

 Spreadsheet::ParseExcel

Optional module File::Log can be used to allow simple logging of errors.


=head1 METHODS

There are no class methods, the object methods are described below.
Private class method start with the underscore character '_' and
should be treated as I<Private>.


=head2 new

Called to create a new BasicReadNamedCol object.  The arguments can
be either a single string (see L<'SYNOPSIS'|"SYNOPSIS">)
which is taken as the filename of the spreadsheet of as named arguments.

 eg.  my $ss = Spreadsheet::BasicReadNamedCol->new(
                  fileName      => 'MyExcelSpreadSheet.xls',
                  skipHeadings  => 1,
                  skipBlankRows => 1,
                  log           => $log,
              );

The following named arguments are available:

=over 4

=item skipHeadings

Don't output the headings line in the first call to
L<'getNextRow'|"getNextRow"> if true.


=item skipBlankRows

Skip blank lines in the spreadsheet if true.


=item log

Use the File::Log object to log exceptions.
If not provided error conditions are logged to STDERR


=item fileName

The name (and optionally path) of the spreadsheet file to process.

=back

=head2 getNextRow()

Get the next row of data from the spreadsheet.  The data is
returned as an array reference.

 eg.  $rowDataArrayRef = $ss->getNextRow();


=head2 numSheets()

Returns the number of sheets in the spreadsheet


=head2 openSpreadsheet(fileName)

Open a new spreadsheet file and set the current sheet to the first
sheet.  The name and optionally path of the
spreadsheet file is a required argument to this method.


=head2 currentSheetNum()

Returns the current sheet number or undef if there is no current sheet.
L<'setCurrentSheetNum'|"setCurrentSheetNum"> can be called to set the
current sheet.


=head2 currentSheetName()

Return the name of the current sheet or undef if the current sheet is
not defined.  see L<'setCurrentSheetNum'|"setCurrentSheetNum">.


=head2 setCurrentSheetNum(num)

Sets the current sheet to the integer value 'num' passed as the required
argument to this method.  Note that this should not be bigger than
the value returned by L<'numSheets'|"numSheets">.


=head2 getNextSheet()

Returns the next sheet "ssBook" object or undef if there are no more sheets
to process.  If there is no current sheet defined the first sheet
is returned.


=head2 getFirstSheet()

Returns the first sheet "ssBook" object.


=head2 cellValue(row, col)

Returns the value of the cell defined by (row, col)in the current sheet.


=head2 getFirstRow()

Returns the first row of data from the spreadsheet (possibly skipping the
column headings  L<'skipHeadings'|"new">) as an array reference.


=head2 setHeadingRow(rowNumber)

Sets the effective minimum row for the spreadsheet to 'rowNumber', since it
is assumed that the heading is on this row and anything above the heading is
not relavent.

B<Note:> the row (and column) numbers are zero indexed.


=head2 logexp(message)

Logs an exception message (can be a list of strings) using the File::Log
object if it was defined and then calls die message.


=head2 logmsg(debug, message)

If a File::Log object was passed as a named argument L<'new'|"new">) and
if 'debug' (integer value) is equal to or greater than the current debug
Level (see File::Log) then the message is added to the log file.

If a File::Log object was not passed to new then the message is output to
STDERR.


=head1 KNOWN ISSUES

None, however please contact the author at gng@cpan.org should you
find any problems and I will endevour to resolve then as soon as
possible


=head1 SEE ALSO

Spreadsheet:ParseExcel on CPAN does all the hard work, thanks
Kawai Takanori (Hippo2000) kwitknr@cpan.org


=head1 AUTHOR

Greg George, IT Technology Solutions P/L, Australia
Mobile: +61-404-892-159, Email: gng@cpan.org


=head1 LICENSE

Copyright (c) 1999- Greg George. All rights reserved. This
program is free software; you can redistribute it and/or modify it under
the same terms as Perl itself.


=head1 CVS ID

$Id: BasicRead.pm,v 1.1.1.1 2004/07/31 07:45:02 Greg Exp $

=cut

#---< End of File >---#