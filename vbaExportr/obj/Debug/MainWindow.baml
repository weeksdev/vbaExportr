   M S B A M L   `   `   `  ����  
vbaExportr    vbaExportr.MainWindow   5      W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   U QMahApps.Metro, Version=0.12.1.0, Culture=neutral, PublicKeyToken=f4fb5a3c4d1e5b4f;controls,http://metro.mahapps.com/winfx/xaml/controls  5   	     � Title$  Exprtr��5   	   $	��500��6   $	��800��6%   (  "MahApps.Metro.Controls.MetroWindow   ShowIconOnTitleBar .   � Icon$ icon\vbaExportr.ico��6L   .��5      � .+�5   
   #  MahApps.Metro.Controls.Flyout  -   5        $��exportFlyout��5         Header$ Workbook VBA Information��6.   $
��Left=�6P   $	��Top=�6k      Position$ Right=�6�   $	��475��6�      IsModal . .��5      � .+�5	      �� -   5        $��fileNameFld��5	      $��25��6+   $	��400��67   $��
File Name:��6C   $
��Left=�6X   $	��Top=�6s   �� 5
      -   5        $��numberOfProjectsFld��5
      $��25��63   $	��400��6?   $��Project Count:��6K   $
��Left=�6d   $	��Top=�6   $��	0,30,05,0q�6�   �� 5      -   5        $��numberOfModules��5      $��25��6/   $	��400��6;   $��Module Count:��6G   $
��Left=�6_   $	��Top=�6z   $��0,60,0,0q�6�   �� 5      -   5           MetroCircleButtonStyle#���   5      $��50��6J   $��50��6V   $��Right=�6a   $��Bottom=�6}   $��30q�6�     appbar_download#��[ 6�   6�   5      5      �� 5   
   -   5        $��uploadFileFld��5   *   $��29��6   $	��182��6   $
��Left=�6?   $	��Top=�6Z   $��	10,10,0,0q�6r    �� 
IsReadOnly . $��Select a file...��6�   -  'MahApps.Metro.Controls.MetroProgressBar  5   
    � Value$ 0Q�6%   $��Bottom=�6/   $��20��6J   �� 5   
   -   5          SquareButtonStyle#���  5      $��Browse��6=   $
��Left=�6N   $��
214,11,0,0q�6i   $	��Top=�6}   $��75��6�   6�   �� 5   
   -   5        $��includeExcelFile��5      $,��&Include Excel File In Extract Location��6+   $��20��6\   $	��250��6h   $
��Left=�6t   $	��Top=�6�   $��10,50q�6�    d� 	IsChecked$
 TrueY�6�   �� 5   
   -	   5        $��deleteBasFiles��5      $-��'Delete existing extracted (*.bas) files��6)   $��20��6[   $	��250��6g   $
��Left=�6s   $	��Top=�6�   $��10,75q�6�   $
 TrueY�6�   5      5      