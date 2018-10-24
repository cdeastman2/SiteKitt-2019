$Users = (Get-ChildItem C:\Users\*).Name
foreach($User in $users){Get-ChildItem "C:\Users\$user\appdata\local\Google\Chrome\User Data\Default\bookmarks*" | Copy-Item -Destination "C:\Users\$user\Desktop"}

Found on Spiceworks: https://community.spiceworks.com/topic/1857338-backup-chrome-bookmarks?utm_source=copy_paste&utm_campaign=growth