#! perl -w
#
#  xlsgrep.pl
#
#  See POD at the end of this file for documentation

use strict;
use Spreadsheet::BasicRead;
use File::Find;

# The pattern is the first argument, otherwise exit
unless ($ARGV[0])
{
	die "SYNTAX: xlsgrep.pl regex_pattern\n";
}

my $pattern = qr/$ARGV[0]/o;

# Find all the .xls files and check the for any cells that match the pattern
find(\&chkXLS, '.');
exit;

#------------------------------------------------------------------------------
#                              End of Main
#------------------------------------------------------------------------------


sub chkXLS
{
	# Called from find.  We only want to check spreadsheet files
	searchXLS($_, $File::Find::name) if /\.xls$/i;
}


sub searchXLS
{
	my ($xlsFileName, $fullPath) = @_;
	my $ss;

	# Open the spreadsheet ready for reading
	unless ($ss = new Spreadsheet::BasicRead($xlsFileName))
	{
		print STDERR "Could not open '$fullPath': $!";
		return;
	}


	# Starting at the first sheet, process each row at a time
	do
	{
		# Track which row, assume zero indexing
		my $row = 0;
		while (my $data = $ss->getNextRow())
		{
			my $col = 0;
			foreach my $col (@$data)
			{
				next unless $col;

				# Check the cell with our pattern
				if ($col =~ /$pattern/)
				{
					# We have a match, so print out the details
					print "$fullPath : Sheet=",
						$ss->currentSheetName(),
						", Row=$row, Col=$col, Value=$col\n";
				}
				$col++;
			}
			$row++;
		}
	} while ($ss->getNextSheet());
}

__END__

=head1 NAME

xlsgrep.pl - Grep spreadsheet files in the current directory and any subdirectories.

=head1 SYNOPSIS

xlsgrep.pl some_regex_pattern

=head1 DESCRIPTION

xlsgrep utilises the power of perls regular expressions to search every cell, on
every sheet in any spreadsheets files found in the current directory or subdirectories.

There are currently no switches supports.  Sum of the standard grep switches can be
handled using perls regular expression syntax.  The equivalent of the ignore case grep
switch (-i) can be applied to I<pattern> by prefixing with B<(?i)> to give I<(?i)pattern>

=head1 SEE ALSO

perlre, perlrequick and perlretut man pages for regualar expression details.

Spreadsheet::BasicRead and Spreadsheet:ParseExcel on CPAN

=head1 AUTHOR

 Greg George, IT Technology Solutions P/L, Australia
 Mobile: +61-404-892-159, Email: gng@cpan.org

=head1 LICENSE

Copyright (c) 1999- Greg George. All rights reserved. This
program is free software; you can redistribute it and/or modify it under
the same terms as Perl itself.


=head1 CVS ID

$Id: xlsgrep.pl,v 1.3 2004/10/03 04:58:20 Greg Exp $

=head1 CVS LOG

$Log: xlsgrep.pl,v $
Revision 1.3  2004/10/03 04:58:20  Greg
- Test of open of spreadsheet and return if failure

Revision 1.2  2004/10/01 10:59:30  Greg
- Replaced the die with print to STDERR when you can't open a spreadsheet

Revision 1.1  2004/09/30 12:31:26  Greg
- Initial development


=cut

#---< End of File >---#