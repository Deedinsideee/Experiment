   M S B A M L   `   `   `  ����Q  MZedGraph, Version=5.1.7.430, Culture=neutral, PublicKeyToken=02a83cbd123fcd605(clr-namespace:ZedGraph;assembly=ZedGraphZedGraph  5   	    Experiments"   Experiments.ExperimentWindow   5   	   W SPresentationCore, Version=6.0.2.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=6.0.2.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=6.0.2.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=6.0.2.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35^ ZWindowsFormsIntegration, Version=6.0.2.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35H 9http://schemas.microsoft.com/winfx/2006/xaml/presentation      5   	   R NSystem.Xaml, Version=6.0.2.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   8zed(clr-namespace:ZedGraph;assembly=ZedGraph     5   	     � Title$-  'Данные экспериментов��5   	   $	��700��5	   	   $
��1000��5
   	    � WindowStartupLocation$ CenterScreen=�5   	   .��5      �  � RowDefinitions 5   
   �� 5      �� 5       � ColumnDefinitions 5   
   �� 5      $
��219* �6   �� 5      $
��281* �6   .+�5   
   &   System.Windows.Controls.DataGrid  -   5          Name$ DataGrid��5      $��0��5      $��0��5      $��2��5      $��10q�5         CanUserAddRows$ False��5         CanUserDeleteRows$ False��5         
IsReadOnly$
 True��5         HeadersVisibility$
 None=�5      	   	RowHeight$	 25Q�5      7  1System.Windows.Forms.Integration.WindowsFormsHost  5   
   $��0��6   $��1��6(   $��10,10,10,10q�68   
   Child.
 5          ZedGraph.ZedGraphControl  -   5        GraphUp��5    "   5!        5"   
   $��1��6   $��1��6(   $��10,10,10,10q�68   .
 5#        -   5        	GraphDown��5#   "   5$      5%      5&      