   M S B A M L   `   `   `  ���� E  ASAP_Import, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null'clr-namespace:SAP_Import
SAP_Import      SAP_Import.MainWindow   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35^ ZWindowsFormsIntegration, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     *localclr-namespace:SAP_Import     -     � Title$  
SAP Import��$	��540��$
��1000�� � WindowStartupLocation$ CenterScreen=� � 
ResizeMode$ CanResizeWithGrip=� � WindowStyle$
 None=� � ShowInTaskbar .  � WindowState$ Normal=����	#00FFFFFF
������  � AllowsTransparency .  � 	Resources  	  Font1��%'   &	        *�/Fonts/#Core Sans N)�.���� 
������� �� CornerRadius$ 10, 10, 10, 10��$��5q����	#FF6495ED._�� 
������ 	 � RowDefinitions	 �� $��50 ��� $
��100* ��� $
��100* ��� $
��500* ��� $��50 �
 � ColumnDefinitions
 �� $��50 ��� $
��100* ��� $
��100* ��� $
��100* ��� $
��100* ��� $
��100* ��� $��50 ��� $��5 ��� $
��200* �.+��� $��0,0,0,20q����	#FFFFFFFF$��1��$��2��$��5��$��1��$ 5�����	#FFC8C8C8$��2q�._�� 
 �� $	��20* ��� $��2 ��� $	��80* �	 �� $��* ��� $��2 ��� $��* �.+��� -    ��Name$
 Name��$��0��$��4q�
������ $��Center=�$��Center=����	#FFAAAAAA$��20(�$��Segoe UI)�$��Normal"�NAME�� $��0��$��2��$��4q�
������ $��Center=�$��Center=����	#FFAAAAAA$��20(�$��Segoe UI)�$��Normal"�WHS�� $��1��$��3�����	#FFC8C8C8�� $��0��$��1��$��3�����	#FFC8C8C8�� -    ��Name$ BomName��$��2��$��	30,0,30,0q�
������ $��0q�$��Stretch=�$��Stretch=�$��Center=�$��Center=����	#FFAAAAAA$��20(�$��Segoe UI)�$��Normal"� �� CharacterCasing$ Upper=�.���� -   $ 	Warehouse��$��2��$��3��$��	30,0,30,0q�
������ $��0q�$��Stretch=�$��Stretch=�$��Center=�$��Center=����	#FFAAAAAA$��20(�$��Segoe UI)�$��Normal"� �� 	MaxLength$ 3��$ Upper=�.���� $��7��$��0��$��5��$��0,0,0,0q����	#FFCCCCCC���	#FF6495ED$��0q��� $��8��$��0��$��5��$��	0,0,100,0q����	#FFCCCCCC���	#FFEEEEEE$��0q��� $��8��$��0��$��5��$��0q����	#FFCCCCCC���	#FFEEEEEE$ 10��._�]� $��2q�.+��� $	��Top=�.+�]� .+��� $
��Left=�$��OUTPUT CONSOLE��$��18(�$
��Bold"�$��3,5,0,0q�$��Segoe UI)�$��Normal$����	#FF707070�� -    ��Name$ 
ExitBorder��$��Right=�$ 4�����	#FFE04F4F$��2q�$��Right=�$��0,3,3,0q�$	��Top=�$��24��$��24��
������ ._��� -    ��Name$
 Exit��$��X�����	#FFE04F4F
�������$��Arial)�$��0q�  
ExitButton#��[ �� -   $ 	TimeBlock�����	#FF707070$��	3, 5, 0,0q�$��Italic$��� $��1�����	#FF969696$��0,2,0,1q��� -    ��Name$ ScrollViewer����. $��Bottom=�
������ $��Hidden=�.���� -	    ��Name$ ConsoleWindow��$��3,1,0,0q��� $��3��$��1��$��2��$ 10��
����d�$��0q�
����d�$��0,30,0,0q� J� SnapsToDevicePixels .  �� BitmapScalingMode$ HighQuality=� �� EdgeMode$ Unspecified=�&   System.Windows.Media.TextOptions   TextFormattingMode$ Ideal=�._��� $ 5��
�������$��2q�
����d�$��0q� . $ HighQuality=�._��� -
   $ 	DragBlock��$ 3��$��2q�
����d�$��2q� J� 	AllowDrop . ���	#99E9E9E9 . $ HighQuality=�._��� -   $ Filename��$��10q�$��18(�$
��Wrap=�$��Center=�$��Center=����	#FF696969$ Ideal=�$ Unspecified=�$
 Fant=�   TextRenderingMode$ 	ClearType=�Drag .XLSX File�� -   $ Convert_Back��{�.  $��3��$��3��$��15,140,15,110q�$ 3�����	#FFF4F4F4���	#FF696969$��2q� . $ HighQuality=�$ Unspecified=�$ Ideal=�._��� -   $ Convert��
������ $��1q�$��Stretch=�$��Stretch=�$��Stretch=�$��Stretch=�$��0q�.��� $��Stretch=�$��Stretch=�.+��� $��Segoe UI)�$��Normal"�$��14(�$��Center=�$��Center=�	CONVERT� -    �Name$ ProgressBar�� � Maximum$	 100Q�$y�Hidden=�$��Stretch=�$��Stretch=�$��0��$��3��$��5���� -   $ OutBlock��$��3��$��4��$��2�� . ���	#FF696969$��2q����	#99E9E9E9$ 5�� . $ HighQuality=�$ Unspecified=�$ Ideal=�$��0,30,0,0q�._��� -   $ Status��$��10q�$��18(�$
��Wrap=�$��Center=�$��Center=����	#FF696969Status