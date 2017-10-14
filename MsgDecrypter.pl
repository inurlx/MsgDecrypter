#!/usr/bin/perl
use strict;
use warnings;
use Term::ANSIColor;
use Email::Outlook::Message;

print color("red"), "\n\n                         By [ MOHAMMED ADEL ] \n\n", color("reset");
print color("white"), "              This is a sample tool that can Decrypt .msg Files . \n\n", color("reset");
print color("blue"), "              [!] Input Examples:\n";
print "               1- 1.msg\n               * It must be on the same directory !!";
DoItAgain:
print color("yellow"), "\n\n[!] Enter the File Name --> ", color("reset");
my $filename = <STDIN>;
if ($filename =~ /\.msg$/i)
{

  print color("red"), "\n[!] Decrypting ... \n", color("reset");
  sleep(1);
      my $msg = new Email::Outlook::Message $filename;
      my $mime = $msg->to_email_mime;
    open(FILE, ">>DESCRIPTED_MSG.txt");
    print FILE $mime->as_string;
    close (FILE);
    print color("white"), "\n\nDecrypted Information Can Be Found At [DESCRIPTED_MSG.txt] \n\n", color("reset");

}
else
{
print color("red"), "\n\n[*] Wrong Input !";
print "\n[*] Make Sure The File You Trying To Decrypt Is With The Extension [.msg] ", color("reset");
goto DoItAgain;
}
