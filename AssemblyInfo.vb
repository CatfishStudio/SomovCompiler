Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices

' ����� �������� � ������ ��������������� ��������� ������� 
' ���������. ����� �������� �������� � ������, �������������� �������� 
' ���� ���������.

' ��������� �������� ��������� ������

<Assembly: AssemblyTitle("Somov Compiler")> 
<Assembly: AssemblyDescription("�����: ����� ������� ��������")> 
<Assembly: AssemblyCompany("��� ��. �.���� �.�������")> 
<Assembly: AssemblyProduct("Somov Compiler")> 
<Assembly: AssemblyCopyright("����� ������� ��������")> 
<Assembly: AssemblyTrademark("CATFISH STUDIO ����� ������� ��������, ���� http://catfish.ltd.ua")> 
<Assembly: CLSCompliant(True)> 

'��������� GUID ������ ��� ������������� ���������� �����, ���� ����
'������ ����� ������� ��� COM
<Assembly: Guid("F773FE41-1AEE-4EDF-BBDB-BABA26B2E44F")> 

' �������� � ������ ������ �������� �������� ����������:
'
'      �������� ����� ������ 
'      �������������� ����� ������ 
'      ����� ������
'      ����� ��������
'
' ����� ������ �������� ��� ���� ���� ����� ��� ������������ ��� ������� ������ � �������� 
' �������� �� ���������, ��������� ���� '*' ��� �������� ����:

<Assembly: AssemblyVersion("1.0.*")> 
