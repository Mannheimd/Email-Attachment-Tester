   M S B A M L   `   `   `  ����R  NEmail Attachment Tester, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullA%clr-namespace:Email_Attachment_TesterEmail_Attachment_Tester  5   	   (    "Email_Attachment_Tester.MainWindow   5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   7local%clr-namespace:Email_Attachment_Tester     5   	     � Title$  Email Attachment Tester��5   	   $	��350��6)   $	��525��66   .��5	      � -   5        .+�5
   
   �� $��
.msg file:��6   $��10,6,0,0q�6%   $	��Top=�67   $
��Left=�6O   �� 5   
   -   5         ��Name$ msgFilePath��5      $��104,10,10,0q�6'   $	��Top=�6<   $��Stretch=�6T   6�   �� 5   
   $��
.xml file:��6   $��	10,29,0,0q�6%   $	��Top=�68   $
��Left=�6P   �� 5   
   -   5        $ xmlFilePath��5      $��104,33,10,0q�6'   $	��Top=�6<   $��Stretch=�6T   6�   �� 5   
   -   5         ��Name$ startButton��5      $��
Start Test��6&   $��	0,0,10,10q�6;   $��Bottom=�6N   $��Right=�6i   6�   �� 5   
   $��History Queue:��6   $��
10,121,0,0q�6)   $	��Top=�6=   $
��Left=�6U   �� 5   
   -   5        $ historyQueuePath��5      $��104,125,10,0q�6,   $	��Top=�6B   $��Stretch=�6Z   6�   �� 5   
   $��Queue Count:��6   $��
10,144,0,0q�6'   $	��Top=�6;   $
��Left=�6S   �� 5   
   -   5        $ 
queueCount��5      $��104,148,0,0q�6&   $��50��6;   $	��Top=�6G   $
��Left=�6_   6z   �� 5   
   $��Success:��6   $��
10,208,0,0q�6#   $	��Top=�67   $
��Left=�6O   �� 5   
   $��Fail:��6   $��
10,234,0,0q�6    $	��Top=�64   $
��Left=�6L   �� 5   
   -   5         ��Name$ successCount��5      $��104,208,10,0q�6&   $	��Top=�6<   �� 5   
   -   5        $ 	failCount��5      $��104,234,10,0q�6#   $	��Top=�69   �� 5   
   -	   5        $ 
stopButton��5      $��	Stop Test��6%   $��	0,0,67,10q�69   $��Bottom=�6L   $��Right=�6g   {�.  6�   �� 5   
   $��Attach Delay (s):��6   $��
10,167,0,0q�6,   $	��Top=�6@   $
��Left=�6X   �� 5   
   -
   5        $ attachDelay��5      $��104,171,0,0q�6'   $��50��6<   $	��Top=�6H   $
��Left=�6`   6{   �� 5   
   $��Files in pool:��6   $��	10,81,0,0q�6)   $	��Top=�6<   $
��Left=�6T   �� 5   
   -   5        $ filesInPool��5      $��104,81,10,0q�6%   $	��Top=�6:   �� 5   
   -   5        $	��Add��5      $
��Left=�6   $��
104,56,0,0q�6:   $	��Top=�6N   $��75��6f   6q   �� 5   
   -   5        $��Reset��5      $
��Left=�6!   $��
184,56,0,0q�6<   $	��Top=�6P   $��75��6h   6s   5      5      