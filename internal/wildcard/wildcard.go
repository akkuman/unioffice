//
// Copyright 2020 FoxyUtils ehf. All rights reserved.
//
// This is a commercial product and requires a license to operate.
// A trial license can be obtained at https://unidoc.io
//
// DO NOT EDIT: generated by unitwist Go source code obfuscator.
//
// Use of this source code is governed by the UniDoc End User License Agreement
// terms that can be accessed at https://unidoc.io/eula/

package wildcard ;func _bc (_ba ,_fc []rune ,_gg int )int {for len (_fc )> 0{switch _fc [0]{default:if len (_ba )==0{return -1;};if _ba [0]!=_fc [0]{return _bc (_ba [1:],_fc ,_gg +1);};case '?':if len (_ba )==0{return -1;};case '*':if len (_ba )==0{return -1;};_da :=_bc (_ba ,_fc [1:],_gg );if _da !=-1{return _gg ;}else {_da =_bc (_ba [1:],_fc ,_gg );if _da !=-1{return _gg ;}else {return -1;};};};_ba =_ba [1:];_fc =_fc [1:];};return _gg ;};func _gaa (_aea ,_ebd []rune ,_ee bool )bool {for len (_ebd )> 0{switch _ebd [0]{default:if len (_aea )==0||_aea [0]!=_ebd [0]{return false ;};case '?':if len (_aea )==0&&!_ee {return false ;};case '*':return _gaa (_aea ,_ebd [1:],_ee )||(len (_aea )> 0&&_gaa (_aea [1:],_ebd ,_ee ));};_aea =_aea [1:];_ebd =_ebd [1:];};return len (_aea )==0&&len (_ebd )==0;};func Index (pattern ,name string )(_af int ){if pattern ==""||pattern =="\u002a"{return 0;};_dc :=make ([]rune ,0,len (name ));_ea :=make ([]rune ,0,len (pattern ));for _ ,_be :=range name {_dc =append (_dc ,_be );};for _ ,_cg :=range pattern {_ea =append (_ea ,_cg );};return _bc (_dc ,_ea ,0);};func Match (pattern ,name string )(_g bool ){if pattern ==""{return name ==pattern ;};if pattern =="\u002a"{return true ;};_ga :=make ([]rune ,0,len (name ));_e :=make ([]rune ,0,len (pattern ));for _ ,_db :=range name {_ga =append (_ga ,_db );};for _ ,_aa :=range pattern {_e =append (_e ,_aa );};_fb :=false ;return _gaa (_ga ,_e ,_fb );};func MatchSimple (pattern ,name string )bool {if pattern ==""{return name ==pattern ;};if pattern =="\u002a"{return true ;};_b :=make ([]rune ,0,len (name ));_f :=make ([]rune ,0,len (pattern ));for _ ,_dg :=range name {_b =append (_b ,_dg );};for _ ,_df :=range pattern {_f =append (_f ,_df );};_cc :=true ;return _gaa (_b ,_f ,_cc );};