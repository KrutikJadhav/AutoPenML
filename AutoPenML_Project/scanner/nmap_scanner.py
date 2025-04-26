import nmap

def run_nmap_scan(target_ip):
    nm = nmap.PortScanner()
    nm.scan(hosts=target_ip, arguments='-T4 -F')
    scan_result = nm[target_ip]
    return scan_result