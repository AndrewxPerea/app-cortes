
principal = 2000000  # Principal en pesos
tasa_interes_mensual = 0.0197  # Tasa de interés mensual
tiempo_meses = 12  # Tiempo en meses

# Cálculo del interés total
interes_total = principal * tasa_interes_mensual * tiempo_meses

# Cálculo del monto total a pagar
monto_total_a_pagar = principal + interes_total

# Cálculo del pago mensual
pago_mensual = monto_total_a_pagar / tiempo_meses

print(monto_total_a_pagar, pago_mensual)
print(pago_mensual)
print(interes_total)
