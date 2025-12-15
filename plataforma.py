import requests
from pprint import pprint
from bs4 import BeautifulSoup
import pandas as pd
cookies = {
    '_gid': 'GA1.3.220837602.1747325943',
    'twk_idm_key': 't41H5_p_nofg0ZkxzxOwn',
    '_gat_gtag_UA_124026907_1': '1',
    '_ga_FN3Y05D0FF': 'GS2.1.s1747325943$o3$g1$t1747327259$j0$l0$h0',
    '_ga': 'GA1.1.694354494.1742480491',
    'TawkConnectionTime': '0',
    'XSRF-TOKEN': 'eyJpdiI6IlYxVW52dHoxZFdJZlJFaEs2dG5waXc9PSIsInZhbHVlIjoiN2RicWNUZ0t0d3ZCRVhUMWo3SHFvdytaZm9CdGswVDdUUEt1dEdYWXA5RXM2cDZuaFQ0d2dTZll1WFptUFVLUiIsIm1hYyI6IjQ0M2NmYzFjNWM4OGMyNjUxN2Q5ZGEwNzA5ZGZlNmU1MTk5NTQxZWNiOTI3MmE1M2RjMmExMmQwYjI1MDk0MWUifQ%3D%3D',
    'laravel_session': 'eyJpdiI6ImJxaUN1QlI3a1ZYWW1qM2VxbXlDalE9PSIsInZhbHVlIjoiYkJBelpDWWlEQzBtdmQ1c1YrRUF1dlF0aEx0YTFXUWZBclZpOFwvTFJwMG54S1ZsZ1wva3J0VU80VDI0dytWcU9NRmxDV1ZoUVlnOWw3RHJCZngwU3R5cTlJbDRCWnd0VjlGTDc1Qzk4YTVOMGcreXE1TjNzVDlBcVwvVjRWNEhGWTgiLCJtYWMiOiIxMDUyNzFhOGNhODFhNWMwZWVjYzRhMDliMWUzZDg0ODY3MmI0YTFhM2I3YTg3NTYxNjBmNTQ5YzUxY2YwNWEwIn0%3D',
}

headers = {
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
    'Accept-Language': 'es-419,es;q=0.9',
    'Cache-Control': 'max-age=0',
    'Connection': 'keep-alive',
    'Referer': 'https://sistegra.com.co/userE',
    'Sec-Fetch-Dest': 'document',
    'Sec-Fetch-Mode': 'navigate',
    'Sec-Fetch-Site': 'same-origin',
    'Sec-Fetch-User': '?1',
    'Upgrade-Insecure-Requests': '1',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36',
    'sec-ch-ua': '"Chromium";v="136", "Google Chrome";v="136", "Not.A/Brand";v="99"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
    # 'Cookie': '_gid=GA1.3.220837602.1747325943; twk_idm_key=t41H5_p_nofg0ZkxzxOwn; _gat_gtag_UA_124026907_1=1; _ga_FN3Y05D0FF=GS2.1.s1747325943$o3$g1$t1747327259$j0$l0$h0; _ga=GA1.1.694354494.1742480491; TawkConnectionTime=0; XSRF-TOKEN=eyJpdiI6IlYxVW52dHoxZFdJZlJFaEs2dG5waXc9PSIsInZhbHVlIjoiN2RicWNUZ0t0d3ZCRVhUMWo3SHFvdytaZm9CdGswVDdUUEt1dEdYWXA5RXM2cDZuaFQ0d2dTZll1WFptUFVLUiIsIm1hYyI6IjQ0M2NmYzFjNWM4OGMyNjUxN2Q5ZGEwNzA5ZGZlNmU1MTk5NTQxZWNiOTI3MmE1M2RjMmExMmQwYjI1MDk0MWUifQ%3D%3D; laravel_session=eyJpdiI6ImJxaUN1QlI3a1ZYWW1qM2VxbXlDalE9PSIsInZhbHVlIjoiYkJBelpDWWlEQzBtdmQ1c1YrRUF1dlF0aEx0YTFXUWZBclZpOFwvTFJwMG54S1ZsZ1wva3J0VU80VDI0dytWcU9NRmxDV1ZoUVlnOWw3RHJCZngwU3R5cTlJbDRCWnd0VjlGTDc1Qzk4YTVOMGcreXE1TjNzVDlBcVwvVjRWNEhGWTgiLCJtYWMiOiIxMDUyNzFhOGNhODFhNWMwZWVjYzRhMDliMWUzZDg0ODY3MmI0YTFhM2I3YTg3NTYxNjBmNTQ5YzUxY2YwNWEwIn0%3D',
}

response = requests.get('https://sistegra.com.co/userN', cookies=cookies, headers=headers)
soup=BeautifulSoup(response.content,"html.parser")
#print(soup.prettify)
tables = soup.find_all('table')
todos_los_clientes=[]
for i, table in enumerate(tables, 1):
    
    rows = table.find_all('tr')
    
    for row in rows:
        cells = row.find_all('td')
        
        row_data = [cell.get_text(strip=True) for cell in cells]
        todos_los_clientes.append(row_data)

#print(todos_los_clientes[1])




import requests

cookies1 = {
    '_gid': 'GA1.3.1326691680.1748455665',
    '_gat_gtag_UA_124026907_1': '1',
    'twk_idm_key': 'LSno0lEH0j1ugfwqtCCjS',
    '_ga_FN3Y05D0FF': 'GS2.1.s1748463020$o18$g1$t1748463037$j43$l0$h0',
    '_ga': 'GA1.1.694354494.1742480491',
    'TawkConnectionTime': '0',
    'XSRF-TOKEN': 'eyJpdiI6ImxYd3hzXC9pSlRPV29UbW5wMjRqS0x3PT0iLCJ2YWx1ZSI6IkFLVjE3YStybXhmOHRHSjhcL3EzTkpxRGkya0FJUDlBZDhlak53am50ZzVtSjFOc3lkQ1o2UFpVMVwvTFIwZVN2dCIsIm1hYyI6ImMzY2U2Y2Q3OWViMjdhZGU5ZDJmMTI3Y2NhNzY2N2QyODFmYzRhMzFjMzYzNzdjMmQ3NWEzNDVjYTZlZDkyN2YifQ%3D%3D',
    'laravel_session': 'eyJpdiI6Ik01bXJBQ1wvYXJTNUlRNm96OVhcL2FkZz09IiwidmFsdWUiOiJTUlwvdmJ3VVlxYTJRU3BQeDVocWFpa1loXC9zRDM5elRzSytsU3B1NFgxUjNTekc5eFwvaGZ0OWpuWGMrMXRJVkVXRnhTSUtQVTc5N3ZETmludlNrNFljc3RCOXdNK0M0UmhvTlwvNUF1K3lrblZ3R0JZM2RmZ3BDSkxwUGhZZ0ZtakgiLCJtYWMiOiJjY2IxNWY3ZTI5ZDZkNGU3OTNmYjQzNTE0ZDIxMzEzMzAzZDg3YWE3N2QxMzk1NDE0Zjg2OTI0NmFlYzAxZDhiIn0%3D',
}

headers1 = {
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
    'Accept-Language': 'es-419,es;q=0.9',
    'Cache-Control': 'max-age=0',
    'Connection': 'keep-alive',
    'Referer': 'https://sistegra.com.co/contrato',
    'Sec-Fetch-Dest': 'document',
    'Sec-Fetch-Mode': 'navigate',
    'Sec-Fetch-Site': 'same-origin',
    'Sec-Fetch-User': '?1',
    'Upgrade-Insecure-Requests': '1',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/137.0.0.0 Safari/537.36',
    'sec-ch-ua': '"Google Chrome";v="137", "Chromium";v="137", "Not/A)Brand";v="24"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
    # 'Cookie': '_gid=GA1.3.1326691680.1748455665; _gat_gtag_UA_124026907_1=1; twk_idm_key=LSno0lEH0j1ugfwqtCCjS; _ga_FN3Y05D0FF=GS2.1.s1748463020$o18$g1$t1748463037$j43$l0$h0; _ga=GA1.1.694354494.1742480491; TawkConnectionTime=0; XSRF-TOKEN=eyJpdiI6ImxYd3hzXC9pSlRPV29UbW5wMjRqS0x3PT0iLCJ2YWx1ZSI6IkFLVjE3YStybXhmOHRHSjhcL3EzTkpxRGkya0FJUDlBZDhlak53am50ZzVtSjFOc3lkQ1o2UFpVMVwvTFIwZVN2dCIsIm1hYyI6ImMzY2U2Y2Q3OWViMjdhZGU5ZDJmMTI3Y2NhNzY2N2QyODFmYzRhMzFjMzYzNzdjMmQ3NWEzNDVjYTZlZDkyN2YifQ%3D%3D; laravel_session=eyJpdiI6Ik01bXJBQ1wvYXJTNUlRNm96OVhcL2FkZz09IiwidmFsdWUiOiJTUlwvdmJ3VVlxYTJRU3BQeDVocWFpa1loXC9zRDM5elRzSytsU3B1NFgxUjNTekc5eFwvaGZ0OWpuWGMrMXRJVkVXRnhTSUtQVTc5N3ZETmludlNrNFljc3RCOXdNK0M0UmhvTlwvNUF1K3lrblZ3R0JZM2RmZ3BDSkxwUGhZZ0ZtakgiLCJtYWMiOiJjY2IxNWY3ZTI5ZDZkNGU3OTNmYjQzNTE0ZDIxMzEzMzAzZDg3YWE3N2QxMzk1NDE0Zjg2OTI0NmFlYzAxZDhiIn0%3D',
}


lista_id_mensuales=[
    134, 3, 8, 10, 18, 20, 26, 27, 33, 34, 37, 40, 43, 49, 50, 52, 53, 56, 57, 60,
    63, 65, 66, 70, 71, 80, 81, 85, 88, 95, 118, 103, 104, 108, 111, 116, 121, 127,
    138, 383, 384, 385, 386, 389, 396, 406, 409, 410, 421, 423, 425, 426, 431, 432,
    436, 437, 439, 450, 451, 452, 466, 470, 475, 480, 485, 487, 491, 509, 510, 515,
    516, 517, 520, 524, 532, 533, 538, 544, 546, 549, 550, 552, 553, 557, 573, 589,
    598, 601, 612, 616, 619, 626, 636, 643, 646, 654, 656, 662, 673, 679, 685, 693,
    703, 705, 712, 713, 716, 721, 724, 727, 730, 731, 734, 736, 738, 740, 741, 744,
    746, 747, 751, 752, 756, 759, 760, 765, 767, 769, 771, 779, 780, 782, 788, 792,
    794, 795, 797, 801, 802, 804, 808, 809, 810, 813, 814, 816, 818, 819, 820, 821,
    824, 828, 829, 833, 834, 837, 838, 839, 843, 844, 852, 854, 855, 862, 863, 866,
    870, 871, 872, 881, 885, 887, 888, 897, 901, 902, 909, 910, 914, 923, 924, 926,
    935, 936, 944, 947, 949, 951, 952, 973, 978, 986, 994, 1009, 1012, 1034, 1035,
    1040, 1043, 1044, 1048, 1053, 1056, 1059, 1060, 1066, 1069, 1073, 1075, 1079,
    1083, 1087, 1088, 1090, 1092, 1094, 1098, 1105, 1121, 1123, 1131, 1141, 1145,
    1154, 1161, 1170, 1173, 1177, 1178, 1182, 1190, 1194, 1199, 1213, 1216, 1218,
    1224, 1229, 1231, 1233, 1235, 1237, 1253, 1275, 1281, 1282, 1292, 1296, 1298,
    1302, 1307, 1309, 1311, 1315, 1326, 1329, 1333, 1335, 1337, 1344, 1346, 1348,
    1352, 1354, 1356, 1358, 1360, 1374, 1376, 1378, 1380, 1382, 1390, 1396, 1398,
    1403, 1406, 1409, 1411, 1413, 1422, 1424, 1426, 1428, 1431, 1433, 1437, 1439,
    1443, 1445, 1447, 1449, 1451, 1453, 1455, 1467, 1469, 1477, 1479, 1486, 1490,
    1500, 1502, 1505, 1507, 1509, 1516, 1518, 1521, 1523, 1525, 1527, 1531, 1533,
    1535, 1537, 1539, 1544, 1549, 1555, 1559, 1565, 1572, 1579, 1581, 1583, 1584,
    1591, 1593, 1597, 1602, 1604, 1610, 1612, 1614, 1616, 1618, 1620, 1626, 1630,
    1632, 1637, 1639, 1641, 1643, 1647, 1650, 1653, 1655, 1658, 1664, 1669, 1679,
    1681, 1683, 1684, 1685, 1688, 1697, 1706, 1720, 1722, 1731, 1735, 1737, 1741,
    1745, 1751, 1759, 1763, 1765, 1770, 1772, 1778, 1780, 1786, 1799, 1803, 1805,
    1811, 1815, 1817, 1821, 1836, 1838, 1840, 1843, 1845, 1847, 1852, 1854, 1856,
    1858, 1860, 1862, 1864, 1869, 1871, 1873, 1875, 1880, 1885, 1889, 1893, 1894,
    1898, 1900, 1902, 1903, 1905, 1908, 1916, 1917, 1918, 1923, 1927, 1930, 1932,
    1934, 1937, 1939, 1945, 1946, 1954, 1959, 1963, 1968, 1971, 1972, 1976, 1977,
    1982, 1984, 1993, 1995, 1998, 2001, 2004, 2007, 2009, 2012, 2015, 2017, 2021,
    2023, 2029, 2031, 2033, 2035, 2036, 2039, 2043, 2046, 2048, 2050, 2053, 2054,
    2062, 2063, 2065, 2068, 2069, 2070, 2071, 2073, 2074, 2075, 2077, 2078, 2081,
    2082, 2083, 2090, 2091, 2092, 2093, 2103, 2104, 2105, 2106, 2107, 2112, 2119,
    2120, 2121, 2122, 2123, 2124, 2125, 2126, 2127, 2130, 2131, 2133, 2134, 2135,
    2137, 2140, 2142, 2143, 2144, 2145, 2146, 2147, 2148, 2149, 2161, 2162, 2167,
    2169, 2171, 2173, 2178, 2183, 2187, 2189, 2191, 2193, 2198, 2201, 2206, 2208,
    2211, 2216, 2218, 2220, 2222, 2226, 2228, 2235, 2239, 2240, 2241, 2243, 2245,
    2251, 2253, 2258, 2260, 2263, 2265, 2266, 2271, 2273, 2275, 2285, 2287, 2292,
    2294, 2298, 2299, 2303, 2306, 2307, 2309, 2311, 2313, 2315, 2317, 2319, 2320,
    2322, 2323, 2324, 2326, 2330, 2332, 2334, 2337, 2341, 2343, 2345, 2346, 2349,
    2353, 2355, 2359, 2363, 2365, 2369, 2375, 2377, 2379, 2384, 2386, 2387, 2388,
    2389, 2390, 2391, 2392, 2393, 2394, 2396, 2397, 2398, 2399, 2400, 2401, 2402,
    2403, 2405, 2406, 2408, 2409, 2410, 2411, 2413, 2414, 2415, 2416, 2417, 2419,
    2420, 2421, 2422, 2423, 2424, 2425, 2426, 2428, 2430, 2431, 2442, 2449, 2455,
    2460, 2473, 2477, 2482, 2486, 2487, 2489, 2490, 2491, 2494, 2496, 2500, 2501,
    2502, 2503, 2504, 2505, 2509, 2511, 2514, 2515, 2516, 2517, 2519, 2523, 2527,
    2529, 2551, 2552, 2554, 2557, 2559, 2560, 2561, 2572, 2574, 2575, 2578, 2580,
    2589, 2590, 2596, 2603, 2604, 2605, 2606, 2607, 2609, 2613, 2614, 2617, 2618,
    2619, 2620, 2621, 2623, 2624, 2626, 2627, 2631, 2632, 2633, 2634, 2636, 2637,
    2638, 2639, 2640, 2641, 2643, 2644, 2645, 2646, 2647, 2648, 2649, 2650, 2652,
    2653, 2654, 2655, 2656, 2658, 2659, 2660, 2661, 2662, 2663, 2664, 2666, 2667
]
#lista_id_mensuales=[2666, 2667,27]
lista_id_anuales=[]
tabla_valores=[]
for elemento in lista_id_mensuales:
    numero=str(elemento)
    url=f'https://sistegra.com.co/contrato/{numero}/edit'
    response1 = requests.get(url, cookies=cookies1, headers=headers1)
    soup1=BeautifulSoup(response1.content,"html.parser")
    #print(soup1.prettify)

    inputs = soup1.find_all('input')
    #clientes_dict=[]
    elements=[]
    lista_id_valido=['razonSocial',"nitRazon","representante",'horasContrato','valorContrato', 'numEmpleados','horasContratoMes']
    for  input_tag in inputs:
        #print(input_tag.get("id",""),"input")
        if input_tag.get("id","") in lista_id_valido:
            elements.append(input_tag.get('value', '') or "")
    labels=soup1.find_all('select')
    clientes_informacion=[]
    lista_id_valido_labels=["idCiudad","tipoVenta","tipoService","nivelRiesgo","cat_riesgos"]
    for label_tag in labels:
        #print(label_tag.get("id",""),"option")
        if label_tag.get("id", "") in lista_id_valido_labels:
            select_id = label_tag.get("id")
            selected_value = label_tag.get('value')
            
            if selected_value:
                selected_option = label_tag.find('option', {'value': selected_value})
                if selected_option:
                    clientes_informacion.append(selected_option.text.strip())
                else:
                    clientes_informacion.append(selected_value)  
            else:
                first_option = label_tag.find('option')
                if first_option:
                    clientes_informacion.append(first_option.text.strip())
                else:
                    clientes_informacion.append('')  
    tabla_valores.append(elements+clientes_informacion)
#print(tabla_valores)
"""for  input_tag in inputs:
    if input_tag.get("id","") in lista_id_valido:
        input_data={
            "name": input_tag.get('name', ''),
            "value": input_tag.get('value', ''),
            "id": input_tag.get('id', '')
        }
        clientes_dict.append(input_data)"""
columns=["Empresa","nit","Representante","#Empleados","Valor contrato","#Conexiones","#Conexiones/mes","Ciudad","Sector económico","Categoria de riesgo","Tipo de servicio","Tipo de venta"]
df=pd.DataFrame(tabla_valores,columns=columns)
df.to_excel("DB_Empresas.xlsx",index=False)

#print(len(clientes_dict))
#print(clientes_dict_labels)
#print(todos_los_clientes)