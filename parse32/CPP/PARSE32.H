
////////////////////////////////////////////////////////////////////
//  FUNCIONES EN LA LIBRERÍA PARSE32.DLL
////////////////////////////////////////////////////////////////////

#ifndef P_EXTERN

#define P_EXTERN(tipo)  extern "C" tipo WINAPI 

#define PHANDLE void*


P_EXTERN(PHANDLE) ParseInit();
P_EXTERN(void) ParseEnd(PHANDLE);
P_EXTERN(void) ParseReset(PHANDLE, int flags=0);

P_EXTERN(int)  GetError(PHANDLE,char* buffer);
P_EXTERN(void) DelValue(PHANDLE,char* nombre);
P_EXTERN(BOOL) DefValue(PHANDLE,char* nombre);

P_EXTERN(BOOL) GetValue(PHANDLE,char* expresion,double *resultado);

P_EXTERN(void) SetValue(PHANDLE,char* nombre,double d);
P_EXTERN(void) SetValueTxt(PHANDLE,char* nombre,char* v);
P_EXTERN(void) SetValueCte(PHANDLE,char* nombre,double d);


#endif // P_EXTERN
