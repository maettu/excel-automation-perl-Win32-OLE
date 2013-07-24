#!/usr/bin/perl
use Modern::Perl;
use Method::Signatures;
use Try::Tiny;
# small interactive shell to Excel.
# supported: insert rows, cols
# write to cells

# TODO
# - record actions to file
# - play recorded actions from file
# - add more actions
# - allow to move active cell
# - "drag-copy" cells

# - h

use FindBin;
use Win32::OLE;

my $file_name = $ARGV[0] or say "usage: mep.pl excel-file" and exit;

# open excel
# TODO open if not already open, otherwise access running instance
my $excel = Win32::OLE->GetActiveObject('Excel.Application') or die $!;

# make Excel visible so we can see what we do.
$excel->{Visible} = 1;


my $book;
if (-f $file_name){
    $book = $excel->Workbooks->Open($file_name);
}
else{
    $book = $excel->Workbooks->Add;
    $book->SaveAs($file_name);
}

my %actions_available = (
    q    => \&exit,
    x    => \&exit,
    exit => \&exit,
    ir   => \&insert_row,
    ic   => \&insert_col,

    t    => \&test,

);

# TODO make this fancier
while (chomp (my $input = <STDIN>)) {
    # the first word becomes the action the rest is arguments.
    my ($action, @arguments) = split (/\s/, $input);

    # direct insert into specified cell. Don't know how to get regex into
    # dispatch table.
    if ($action =~ /^\w+\d+$/){
        my $content = join " ", @arguments;
        say $content;
        $book->ActiveSheet->Range($action)->{Value} = $content;
        next;
    }

    try{
        no warnings;
        $actions_available{$action}->(@arguments);
    }
    catch{
        say "error: $_";
        say "$action unsupported, h for help";
    };


}

# test here whatever you need tested. To be removed.
func test(@args){
    say $book->ActiveSheet->ActiveRow;
}

sub exit{
    $book->Save();
    $book->Close();
    say "goodbye" and exit;
}

# TODO: insert at active row / col. Problem: how to get active row/col..
func insert_col($col//='A', $amount//=1){
    foreach (1 .. $amount){
        $book->ActiveSheet->Columns("$col:$col")->Insert();
    }
}

func insert_row($row//=1, $amount//=1){
    foreach (1 .. $amount){
        $book->ActiveSheet->Rows("$row:$row")->Insert();

        # this is supposed to work but..
        # $curSheet->Rows("2:2")->Insert({Shift => xlDown});

        # alternative implementation, left here for reference
#~         $book->ActiveSheet->Cells($row,1)->EntireRow->Insert;
    }
}
