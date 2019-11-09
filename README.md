# xmax_gift
## Intro
Simple program for randomly extracting a friend to which make a gift (for XMAS!)

We are a large group of friends, and we want to have xmas gift!
But...how to do it? Many many different gift to be done by everybody!!

So we decided to make an extraction: everyone has to make a gift to a specific random friend;
and in order to respect some constraint (relative don't want to extract their sister, or husband, ...)
Why not to do it in Python?


## Tutorial
How to setup the program, and which file are needed.


#### Mail
file named "xmas_conf.csv" is needed to give the mail and password of the sender (and of who will receive the summary).
The mail should be gmail, and the use from application need to be granted
the file will be like this:
```
mail;mymail@gmail.com
pass;mypassword
```

#### List of partecipant
An excel file with the list of partecipant should be prepared; the filename is not important, because you will manually select the file with UI.
The structure of the file is:
```
who_give	mail	                exclusion
Marco	    mail1@hotmail.it	    Marta
Marta	    mail2@hotmail.it	    
Stefy	    mail3@gmail.com	    Alex, Marta, Paolo
...
```


#### Go!
Launch the program, and every partecipant will receive a mail with the name of the extracted person to which he needs to prepare and give a present!
The sender mail will also receive a summary of the extraction, that will also be saved in excel file in the folder!

Happy Christmas