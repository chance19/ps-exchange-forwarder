# Before running this script please update the import.csv file with relevant email addresses

# Import data from csv
$users = import-csv -path import.csv

# For each entry in the csv create a new mail contact in active directory, in a specified OU and then set the exchange mailbox to deliver to that contact.
foreach ($u in $users){

    New-MailContact -Name $u.name -ExternalEmailAddress $u.fwd -OrganizationalUnit "OU=MailForwarders,DC=company,DC=com"
    Set-Mailbox $u.eml -ForwardingAddress $u.fwd -DeliverToMailboxAndForward $false

}