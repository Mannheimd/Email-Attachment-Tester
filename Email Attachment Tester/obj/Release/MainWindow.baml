   M S B A M L   `   `   `  ���� R  NEmail Attachment Tester, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullA%clr-namespace:Email_Attachment_TesterEmail_Attachment_Tester  (    "Email_Attachment_Tester.MainWindow   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     7local%clr-namespace:Email_Attachment_Tester       � Title$  Email Attachment Tester��$	��350��$	��525��.��� -   .+��� $��
.msg file:��$��10,6,0,0q�$	��Top=�$
��Left=��� -    ��Name$ msgFilePath��$��104,10,10,0q�$	��Top=�$��Stretch=��� $��
.xml file:��$��	10,29,0,0q�$	��Top=�$
��Left=��� -   $ xmlFilePath��$��104,33,10,0q�$	��Top=�$��Stretch=��� -    ��Name$ startButton��$��
Start Test��$��	0,0,10,10q�$��Bottom=�$��Right=��� $��History Queue:��$��
10,121,0,0q�$	��Top=�$
��Left=��� -   $ historyQueuePath��$��104,125,10,0q�$	��Top=�$��Stretch=��� $��Queue Count:��$��
10,144,0,0q�$	��Top=�$
��Left=��� -   $ 
queueCount��$��104,148,0,0q�$��50��$	��Top=�$
��Left=��� $��Success:��$��
10,208,0,0q�$	��Top=�$
��Left=��� $��Fail:��$��
10,234,0,0q�$	��Top=�$
��Left=��� -    ��Name$ successCount��$��104,208,10,0q�$	��Top=��� -   $ 	failCount��$��104,234,10,0q�$	��Top=��� -	   $ 
stopButton��$��	Stop Test��$��	0,0,67,10q�$��Bottom=�$��Right=�{�.  �� $��Attach Delay (s):��$��
10,167,0,0q�$	��Top=�$
��Left=��� -
   $ attachDelay��$��104,171,0,0q�$��50��$	��Top=�$
��Left=��� $��Files in pool:��$��	10,81,0,0q�$	��Top=�$
��Left=��� -   $ filesInPool��$��104,81,10,0q�$	��Top=��� -   $	��Add��$
��Left=�$��
104,56,0,0q�$	��Top=�$��75���� -   $��Reset��$
��Left=�$��
184,56,0,0q�$	��Top=�$��75��