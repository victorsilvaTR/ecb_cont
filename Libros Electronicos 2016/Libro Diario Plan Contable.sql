USE SAFC_ECB
GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
/*-----------------------------------------------------------------------------------------------------------------                                                                                          
MODULO DE CONTABILIDAD                                                                                          
Creador: Pool Berrospi                                                          
Fecha Creaciòn: 28/04/2014              
DESCRIPCION  : Reporte de Libro Diario Electronico de Detalle de Cuentas                                                      
------------------------------------------------------------------------------------------------------------------*/                                                                                          
--DiarioDetalleElectronico '042','2014','04','LD'              
                     
CREATE PROCEDURE [dbo].[DiarioDetalleElectronico1]              
 @Emp_cCodigo char(3)='',              
 @Pan_cAnio char(4)='',              
 @Per_cPeriodo char(2)='',              
 @Lib_cTipoLibro char(2)=''              
 --WITH ENCRYPTION    
As    
Declare @Separador varchar(1)    
Declare @NumReg int    
Declare @Periodo varchar(2)    
Declare @fecha varchar(10)    
    
set dateformat dmy    
    
set @fecha = '01/' + @Per_cPeriodo + '/'+ @Pan_cAnio    
    
Set @Separador = '|'    
    
select @NumReg = COUNT(*), @Periodo = Per_cPeriodo from CNT_lIBROSGENERADOS    
where Emp_cCodigo = @Emp_cCodigo and Pan_cAnio =@Pan_cAnio and Lib_cTipoLibro = @Lib_cTipoLibro    
group by Per_cPeriodo    
    
if @@ROWCOUNT = 0    
begin    
    
set @fecha = day(DATEadd(D,-(day(dateadd(M,1,cast(@fecha as datetime)))),dateadd(M,1,cast(@fecha as datetime))))    
             
--select  (LTRIM(RTRIM(LEFT(@Pan_cAnio + @Per_cPeriodo + @fecha   ,8))) + @Separador +              
--Pla_cCuentaContable + @Separador +              
--LTRIM(RTRIM(LEFT(replace(replace(replace(Pla_cNombreCuenta, char(13), ' '), char(10), ' '), char(9), ' '),100))) + @Separador +       
--'01'   + @Separador +              
--'-' + @Separador +              
--Pla_dEstadoO + @Separador) as 'Registro'              
--from CNM_PLAN_CTA    
--where Emp_cCodigo = @Emp_cCodigo and Pan_cAnio =@Pan_cAnio and Pla_cTitulo = 'N'      

DECLARE @TipoPlan CHAR(1)
SELECT @TipoPlan = CCL.Cfl_cTipoPlan FROM dbo.CNT_CONFIG_LIBROS CCL
WHERE CCL.Emp_cCodigo = @Emp_cCodigo AND CCL.Pan_cAnio = @Pan_cAnio AND CCL.Cfl_cDeleted <> '*'

SELECT (LTRIM(RTRIM(LEFT(@Pan_cAnio + @Per_cPeriodo + @fecha   ,8))) + @Separador + 
        REPLACE(REPLACE(LEFT(cpc.Pla_cCuentaContable, 100), ' / ', ''), 'S/.', 'S.') + @Separador + 
        cpc.Pla_cNombreCuenta + @Separador + 
        CASE WHEN @TipoPlan = '0' THEN '02' ELSE '01' END + @Separador + 
        '-' + @Separador + 
        '' + @Separador + 
        '' + @Separador + 
        cpc.Pla_dEstadoO + @Separador) AS Registro  FROM dbo.CNM_PLAN_CTA CPC
WHERE CPC.Emp_cCodigo = @Emp_cCodigo AND CPC.Pan_cAnio = @Pan_cAnio AND CPC.Pla_cDeleted <> '*' AND CPC.Pla_cTitulo = 'N'
    
    
end    
else    
begin    
    
select top 1 @Periodo = Per_cPeriodo from CNT_lIBROSGENERADOS    
where Emp_cCodigo = @Emp_cCodigo and Pan_cAnio =@Pan_cAnio and Lib_cTipoLibro = @Lib_cTipoLibro    
Order by Per_cPeriodo  
              
select  (LTRIM(RTRIM(LEFT(@Pan_cAnio +  @Periodo + convert(char(2),day(DATEadd(D,-(day(dateadd(M,1,Pla_dFechaModifica))),dateadd(M,1,Pla_dFechaModifica)))),8))) + @Separador +              
Pla_cCuentaContable + @Separador +  
LTRIM(RTRIM(LEFT(replace(replace(replace(Pla_cNombreCuenta, char(13), ' '), char(10), ' '), char(9), ' '),100))) + @Separador +         
'01'   + @Separador +              
'-' + @Separador +    
'' + @Separador + 
'' + @Separador +          
Pla_dEstadoD + @Separador) as 'Registro'              
from CNM_PLAN_CTA               
where Emp_cCodigo = @Emp_cCodigo and Pan_cAnio =@Pan_cAnio and Pla_dEstadoD in ('8','9')              
and month(Pla_dFechaModifica) = @Per_cPeriodo and year(Pla_dFechaModifica)= @Pan_cAnio and Pla_cTitulo = 'N'        
              
end  
GO
