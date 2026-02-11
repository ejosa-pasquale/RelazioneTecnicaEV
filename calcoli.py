from __future__ import annotations

from dataclasses import dataclass
from math import sqrt
from typing import Dict, Tuple, Optional

# Tabelle semplificate R e X (ohm/km) per rame, isolante PVC 70°C (valori tipici/indicativi).
# Per uso "relazione DiCo" (verifiche di massima). In contesti critici usare dati di catalogo/cavi reali.
RX_CU_70C_OHM_KM: Dict[float, Tuple[float, float]] = {
    1.5: (12.10, 0.080),
    2.5: (7.41, 0.075),
    4:   (4.61, 0.070),
    6:   (3.08, 0.068),
    10:  (1.83, 0.065),
    16:  (1.15, 0.062),
    25:  (0.727,0.060),
    35:  (0.524,0.058),
    50:  (0.387,0.056),
    70:  (0.268,0.054),
    95:  (0.193,0.053),
    120: (0.153,0.052),
    150: (0.124,0.051),
}

@dataclass
class DropResult:
    delta_v_volt: float
    delta_v_percent: float

def corrente_da_potenza(p_kw: float, alimentazione: str, cosphi: float = 0.95, v_ll: float = 400.0, v_ln: float = 230.0) -> float:
    """
    Restituisce la corrente di impiego Ib [A] a partire dalla potenza attiva [kW].
    alimentazione: 'Monofase 230 V' oppure 'Trifase 400 V'
    """
    p_w = p_kw * 1000.0
    if alimentazione.lower().startswith("mono"):
        return p_w / (v_ln * max(cosphi, 0.1))
    return p_w / (sqrt(3) * v_ll * max(cosphi, 0.1))

def caduta_tensione(
    i_a: float,
    l_m: float,
    sezione_mm2: float,
    alimentazione: str,
    cosphi: float = 0.95,
    sinphi: Optional[float] = None,
) -> DropResult:
    """
    Calcolo ΔV con R e X tabellari (ohm/km). L in metri (percorso).
    - Monofase: ΔV = 2 * I * (R cosφ + X sinφ) * L[km]
    - Trifase:  ΔV = √3 * I * (R cosφ + X sinφ) * L[km]
    """
    if sinphi is None:
        sinphi = sqrt(max(0.0, 1.0 - cosphi**2))
    if sezione_mm2 not in RX_CU_70C_OHM_KM:
        # fallback: interpolazione semplice tra sezioni note
        keys = sorted(RX_CU_70C_OHM_KM.keys())
        # clamp
        if sezione_mm2 <= keys[0]:
            r, x = RX_CU_70C_OHM_KM[keys[0]]
        elif sezione_mm2 >= keys[-1]:
            r, x = RX_CU_70C_OHM_KM[keys[-1]]
        else:
            # find interval
            for a,b in zip(keys, keys[1:]):
                if a <= sezione_mm2 <= b:
                    ra, xa = RX_CU_70C_OHM_KM[a]
                    rb, xb = RX_CU_70C_OHM_KM[b]
                    t = (sezione_mm2 - a) / (b - a)
                    r = ra + t*(rb-ra)
                    x = xa + t*(xb-xa)
                    break
    else:
        r, x = RX_CU_70C_OHM_KM[sezione_mm2]
    l_km = l_m / 1000.0
    term = (r * cosphi + x * sinphi) * l_km
    if alimentazione.lower().startswith("mono"):
        dv = 2.0 * i_a * term
        v_nom = 230.0
    else:
        dv = sqrt(3) * i_a * term
        v_nom = 400.0
    return DropResult(delta_v_volt=dv, delta_v_percent=(dv / v_nom) * 100.0)

def ia_magnetotermico(curva: str, in_a: float) -> float:
    """
    Corrente di intervento istantaneo indicativa (IEC 60898):
    B ~ 5·In, C ~ 10·In, D ~ 20·In
    """
    curva = (curva or "C").strip().upper()
    mult = {"B": 5.0, "C": 10.0, "D": 20.0}.get(curva, 10.0)
    return mult * max(in_a, 0.1)

def zs_massima_tn(u0: float, curva: str, in_a: float) -> float:
    """
    Zs_max = U0 / Ia (approccio semplificato).
    """
    ia = ia_magnetotermico(curva, in_a)
    return u0 / ia

def verifica_tt_ra_idn(ra_ohm: float, idn_a: float, ul: float = 50.0) -> bool:
    """
    Verifica semplificata TT: Ra * Idn <= UL (50V tipico).
    """
    return (ra_ohm * idn_a) <= ul
