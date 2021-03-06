FasdUAS 1.101.10   ��   ��    k             l     ��  ��    7 1 Get a list of all "current messages" in Outlook.     � 	 	 b   G e t   a   l i s t   o f   a l l   " c u r r e n t   m e s s a g e s "   i n   O u t l o o k .   
  
 l     ����  O         r        l    ����  1    ��
�� 
CMgs��  ��    o      ���� "0 currentmessages currentMessages  5     �� ��
�� 
capp  m       �   * c o m . m i c r o s o f t . O u t l o o k
�� kfrmID  ��  ��        l     ��������  ��  ��        l     ��  ��    !  Loop through the messages.     �   6   L o o p   t h r o u g h   t h e   m e s s a g e s .   ��  l   � ����  X    � ��   k    �      ! " ! O    � # $ # k   ' � % %  & ' & l  ' '�� ( )��   ( ) # Only notify about unread messages.    ) � * * F   O n l y   n o t i f y   a b o u t   u n r e a d   m e s s a g e s . '  +�� + Z   ' � , -�� . , =  ' , / 0 / n   ' * 1 2 1 1   ( *��
�� 
pRed 2 o   ' (���� 0 eachmessage eachMessage 0 m   * +��
�� boovfals - Z   / � 3 4�� 5 3 =  / 4 6 7 6 n   / 2 8 9 8 1   0 2��
�� 
pMfd 9 o   / 0���� 0 eachmessage eachMessage 7 m   2 3��
�� boovfals 4 k   7 � : :  ; < ; r   7 : = > = m   7 8��
�� boovtrue > o      ���� *0 displaynotification displayNotification <  ? @ ? r   ; A A B A e   ; ? C C l  ; ? D���� D n   ; ? E F E 1   < >��
�� 
subj F o   ; <���� 0 eachmessage eachMessage��  ��   B o      ����  0 messagesubject messageSubject @  G H G r   B G I J I n   B E K L K 1   C E��
�� 
sndr L o   B C���� 0 eachmessage eachMessage J o      ���� 0 messagesender messageSender H  M N M r   H Q O P O n   H M Q R Q 1   I M��
�� 
PlTC R o   H I���� 0 eachmessage eachMessage P o      ����  0 messagecontent messageContent N  S T S Z   R y U V���� U ?  R ] W X W n   R Y Y Z Y 1   U Y��
�� 
leng Z o   R U����  0 messagecontent messageContent X m   Y \���� ( V r   ` u [ \ [ n   ` q ] ^ ] 7  c q�� _ `
�� 
ctxt _ m   i k����  ` m   l p���� ( ^ o   ` c����  0 messagecontent messageContent \ o      ����  0 messagecontent messageContent��  ��   T  a b a l  z z�� c d��   c ` Z Get an appropriate representation of the sender; preferably name, but fall back on email.    d � e e �   G e t   a n   a p p r o p r i a t e   r e p r e s e n t a t i o n   o f   t h e   s e n d e r ;   p r e f e r a b l y   n a m e ,   b u t   f a l l   b a c k   o n   e m a i l . b  f�� f Q   z � g h i g Z   } � j k�� l j =  } � m n m n   } � o p o 1   ~ ���
�� 
pnam p o   } ~���� 0 messagesender messageSender n m   � � q q � r r   k r   � � s t s n   � � u v u 1   � ���
�� 
radd v o   � ����� 0 messagesender messageSender t o      ���� 0 messagesender messageSender��   l r   � � w x w n   � � y z y 1   � ���
�� 
pnam z o   � ����� 0 messagesender messageSender x o      ���� 0 messagesender messageSender h R      �� { |
�� .ascrerr ****      � **** { o      ���� 0 errormessage errorMessage | �� }��
�� 
errn } o      ���� 0 errornumber errorNumber��   i Q   � � ~  � ~ r   � � � � � n   � � � � � 1   � ���
�� 
radd � o   � ����� 0 messagesender messageSender � o      ���� 0 messagesender messageSender  R      �� � �
�� .ascrerr ****      � **** � o      ���� 0 errormessage errorMessage � �� ���
�� 
errn � o      ���� 0 errornumber errorNumber��   � k   � � � �  � � � l  � ��� � ���   � H B Couldn�t get name or email; we�ll just say the sender is unknown.    � � � � �   C o u l d n  t   g e t   n a m e   o r   e m a i l ;   w e  l l   j u s t   s a y   t h e   s e n d e r   i s   u n k n o w n . �  ��� � r   � � � � � m   � � � � � � �  U n k n o w n   s e n d e r � o      ���� 0 messagesender messageSender��  ��  ��   5 k   � � � �  � � � l  � ��� � ���   � Z T The message was already marked for deletion, so we won�t bother notifying about it.    � � � � �   T h e   m e s s a g e   w a s   a l r e a d y   m a r k e d   f o r   d e l e t i o n ,   s o   w e   w o n  t   b o t h e r   n o t i f y i n g   a b o u t   i t . �  ��� � r   � � � � � m   � ���
�� boovfals � o      ���� *0 displaynotification displayNotification��  ��   . k   � � � �  � � � l  � ��� � ���   � K E The message was already read, so we won�t bother notifying about it.    � � � � �   T h e   m e s s a g e   w a s   a l r e a d y   r e a d ,   s o   w e   w o n  t   b o t h e r   n o t i f y i n g   a b o u t   i t . �  ��� � r   � � � � � m   � ���
�� boovfals � o      ���� *0 displaynotification displayNotification��  ��   $ 5    $�� ���
�� 
capp � m   ! " � � � � � * c o m . m i c r o s o f t . O u t l o o k
�� kfrmID   "  � � � l  � ���������  ��  ��   �  � � � l  � ��� � ���   �   Display notification    � � � � *   D i s p l a y   n o t i f i c a t i o n �  � � � Z   � � � ����� � =  � � � � � o   � ����� *0 displaynotification displayNotification � m   � ���
�� boovtrue � k   � � � �  � � � l  � ��� � ���   � &   Without using terminal-notifier    � � � � @   W i t h o u t   u s i n g   t e r m i n a l - n o t i f i e r �  � � � l  � ��� � ���   � m g display notification messageContent with title messageSender subtitle messageSubject sound name "Ping"    � � � � �   d i s p l a y   n o t i f i c a t i o n   m e s s a g e C o n t e n t   w i t h   t i t l e   m e s s a g e S e n d e r   s u b t i t l e   m e s s a g e S u b j e c t   s o u n d   n a m e   " P i n g " �  � � � l  � ��� � ���   � 2 , Allow time for the notification to trigger.    � � � � X   A l l o w   t i m e   f o r   t h e   n o t i f i c a t i o n   t o   t r i g g e r . �  � � � l  � ��� � ���   �   delay 1    � � � �    d e l a y   1 �  � � � l  � ���������  ��  ��   �  � � � l  � ��� � ���   � _ Y Show notification using terminal-notifier: https://github.com/julienXX/terminal-notifier    � � � � �   S h o w   n o t i f i c a t i o n   u s i n g   t e r m i n a l - n o t i f i e r :   h t t p s : / / g i t h u b . c o m / j u l i e n X X / t e r m i n a l - n o t i f i e r �  � � � l  � ��� � ���   � O I First install terminal-notifier with: sudo gem install terminal-notifier    � � � � �   F i r s t   i n s t a l l   t e r m i n a l - n o t i f i e r   w i t h :   s u d o   g e m   i n s t a l l   t e r m i n a l - n o t i f i e r �  ��� � I  � ��� ���
�� .sysoexecTEXT���     TEXT � b   � � � � � b   � � � � � b   � � � � � b   � � � � � b   � � � � � b   � � � � � m   � � � � � � � R / u s r / l o c a l / b i n / t e r m i n a l - n o t i f i e r   - t i t l e   ' � o   � ����� 0 messagesender messageSender � m   � � � � � � �  '   - s u b t i t l e   ' � o   � �����  0 messagesubject messageSubject � m   � � � � � � �  '   - m e s s a g e   ' � o   � �����  0 messagecontent messageContent � m   � � � � � � � r '   - s o u n d   ' P i n g '   - e x e c u t e   ' o p e n   - b   c o m . m i c r o s o f t . O u t l o o k '  ��  ��  ��  ��   �  ��� � l  � ���������  ��  ��  ��  �� 0 eachmessage eachMessage  o    ���� "0 currentmessages currentMessages��  ��  ��       �� � ���   � ��
�� .aevtoappnull  �   � **** � �� ����� � ���
�� .aevtoappnull  �   � **** � k     � � �  
 � �  ����  ��  ��   � �������� 0 eachmessage eachMessage�� 0 errormessage errorMessage�� 0 errornumber errorNumber �  �� ����~�}�|�{ ��z�y�x�w�v�u�t�s�r�q�p�o�n q�m�l � � � � � ��k
�� 
capp
�� kfrmID  
� 
CMgs�~ "0 currentmessages currentMessages
�} 
kocl
�| 
cobj
�{ .corecnte****       ****
�z 
pRed
�y 
pMfd�x *0 displaynotification displayNotification
�w 
subj�v  0 messagesubject messageSubject
�u 
sndr�t 0 messagesender messageSender
�s 
PlTC�r  0 messagecontent messageContent
�q 
leng�p (
�o 
ctxt
�n 
pnam
�m 
radd�l 0 errormessage errorMessage � �j�i�h
�j 
errn�i 0 errornumber errorNumber�h  
�k .sysoexecTEXT���     TEXT�� �)���0 *�,E�UO ��[��l kh  )���0 ���,f  ���,f  �eE�O��,EE�O��,E�O�a ,E` O_ a ,a  _ [a \[Zk\Za 2E` Y hO "�a ,a   �a ,E�Y 	�a ,E�W X   �a ,E�W X  a E�Y fE�Y fE�UO�e   a �%a %�%a %_ %a %j Y hOP[OY�+ ascr  ��ޭ