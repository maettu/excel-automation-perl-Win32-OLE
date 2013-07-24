# from http://stackoverflow.com/questions/5399589/how-can-i-programmatically-discover-a-win32ole-objects-properties-and-methods/5404271#5404271


use Win32::OLE;

my $xls = Win32::OLE->new('Excel.Application', 'Quit');

$xls->Workbooks->Add;

foreach my $key (sort keys %$xls){
    my $value;

    eval {$value = $xls->{$key}};
    $value = "***Exception" if $@;
    $value = "<undef>" unless defined $value;

    $value = '['.Win32::OLE->QueryObjectType($vaue).']'
        if UNIVERSAL::isa($value, 'Win32::OLE');

    $value = '('.join(',', @$value).')'
        if ref $value eq "ARRAY";

    printf "%s %s %s\n", $key, '.' x (40-length($key)), $value;

}
