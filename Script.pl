use Win32::OLE;

$object = Win32::OLE->new('PERLVBExample.PerlVBExm');

$s = $object->{NewValue};
print "NewValue property returns: $s\n";
print "Created by: Trenton Miller\n";

$object->{SafeValue} = 5;
$s2 = $object->SafeValue;
print "addem(2, 2) = $s3\n";

$object->ShowForm;