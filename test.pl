# Before `make install' is performed this script should be runnable with
# `make test'. After `make install' it should work as `perl test.pl'

#########################

# change 'tests => 1' to 'tests => last_test_to_print';

use Test;
BEGIN { plan tests => 5 };
use Spreadsheet::BasicRead;
print 'Use it...........................';
ok(1); # If we made it this far, we're ok.

#########################

# Insert your test code below, the Test module is use()ed here so read
# its man page ( perldoc Test ) for help writing this test script.

## Can we create a log object
print 'Create object....................';
my $ss;
ok( sub { $ss = Spreadsheet::BasicRead->new('Test.xls'); } );


## Read from it
my $data = $ss->getNextRow();
print "Reading row, (", scalar(@$data), ") columns.........";
ok( scalar(@$data), 3 );


## Get the sheet name
my $name = $ss->currentSheetName();
print "Getting sheet name ($name).";
ok( $name, "/Test Sheet1/" );

## Get the number of sheets
my $cnt = $ss->numSheets();
print "Getting number sheets ($cnt)........";
ok( $cnt, 3 );

#---<end of File >---#