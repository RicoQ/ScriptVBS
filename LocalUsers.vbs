'on error resume next

     set OWMI=getobject("winmgmts:")
     set les_Users =OWMI.ExecQuery("select * from Win32_UserAccount")

  for each User in les_Users
     wscript.echo User.caption
  next
