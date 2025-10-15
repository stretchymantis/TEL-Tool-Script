
msgbox % GetUserFullName( A_UserName )



GetUserFullName( usr="" )

{

	IfEqual,usr,

		usr=%A_UserName%

	StringLen,L,usr

	varSetCapacity(sUserW, L*2+2, 0)

	dllCall("MultiByteToWideChar","uint",0, "uint",0

				,"str",usr, "uint",-1, "uint",&sUserW, "uint",L+1)



;typedef struct _USER_INFO_2 {

	usri2_nameW=0

	usri2_passwordW=4

	usri2_password_ageL=8

	usri2_privL=12

	usri2_home_dirW=16

	usri2_commentW=20

	usri2_flagsL=24

	usri2_script_pathW=28

	usri2_auth_flagsL=32

	usri2_full_nameW=36

	usri2_usr_commentW=40

	usri2_parmsW=44

	usri2_workstationsW=48

	usri2_last_logonL=52

	usri2_last_logoffL=56

	usri2_acct_expiresL=60

	usri2_max_storageL=64

	usri2_units_per_weekL=68

	usri2_logon_hours=72

	usri2_bad_pw_countL=76

	usri2_num_logonsL=80

	usri2_logon_serverW=84

	usri2_country_codeL=88

	usri2_code_pageL=92



	DllCall("netapi32\NetUserGetInfo", "uint",0, "uint",&sUserW, "uint",2, "uintp",pUsrInfo2)

	pUFN:=NumGet( 0+pUsrInfo2, usri2_full_nameW )



	L:=1+DllCall("lstrlenW","uint",pUFN)

	VarSetCapacity( sRet, L)

	dllCall("WideCharToMultiByte","uint",0, "uint",0, "uint",pUFN

				  , "int",-1, "str",sRet, "uint",L, "uint",0, "uint",0)

	return sRet

}