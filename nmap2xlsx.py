import openpyxl
import argparse

def extract_ip_addresses(filename):
    ip_address_ports = {}

    with open(filename, 'r') as file:
        ip = ""
        for line in file:
            if line.startswith("Nmap scan report for ") and "[host down, received no-response]" not in line and "[host down, received host-unreach]" not in line:
                ip = line.split("Nmap scan report for ")[-1].replace('(', ',').replace(')','').strip()
            elif "open" in line and "Discovered" not in line and "tcp" in line:
                if ip in ip_address_ports:
                    ip_address_ports[ip].append(line.strip())
                else:
                    ip_address_ports[ip] = [line.strip()]

    return ip_address_ports

def write_to_excel(filename, ip_address_ports):
    wb = openpyxl.Workbook()
    ws = wb.active

    ws['A1'] = "Host"
    ws['B1'] = "Port, service"

    row = 2
    for ip, ports in ip_address_ports.items():
        ws.cell(row=row, column=1, value=ip)
        ws.cell(row=row, column=2, value='\n'.join(ports).replace("syn-ack", '').replace('tcp', 'TCP').replace('?', ''))
        row += 1

    wb.save(filename)

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('input_filename')
    parser.add_argument('output_filename')
    args = parser.parse_args()
    input_filename = args.input_filename
    output_filename = args.output_filename

    ip_address_ports = extract_ip_addresses(input_filename)
    write_to_excel(output_filename, ip_address_ports)

    print("Успешно сохранено в файл ", output_filename)
