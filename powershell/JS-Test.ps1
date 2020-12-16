# Пример внедрения кода на JS
Add-Type @'
 class Calc {
   function sum(a,b){
       return a + b;
        }
      }
'@ -Language JScript


$rect = [Calc]::new()
$rect.sum(2, 2)