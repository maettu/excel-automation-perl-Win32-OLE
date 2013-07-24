# from http://stackoverflow.com/questions/5399589/how-can-i-programmatically-discover-a-win32ole-objects-properties-and-methods/5404271#5404271


use Modern::Perl;
use Win32::OLE;

my $OleObject = Win32::OLE->new('Excel.Application', 'Quit');


my $typeinfo = $OleObject->GetTypeInfo();
my $attr = $typeinfo->_GetTypeAttr();
for (my $i = 0; $i< $attr->{cFuncs}; $i++) {
    my $desc = $typeinfo->_GetFuncDesc($i);
    # the call conversion of method was detailed in %$desc
    my $funcname = @{$typeinfo->_GetNames($desc->{memid}, 1)}[0];
    say $funcname;
}
