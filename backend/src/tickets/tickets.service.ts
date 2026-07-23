import { Injectable } from '@nestjs/common';

export interface TicketDTO {
  id: string;
  matchTitle: string;
  buyerName: string;
  tribune: string;
  seatCode: string;
  price: number;
  qrHash: string;
  status: 'VALIDE' | 'SCANNE' | 'CANCELLED';
  scannedAt?: string;
}

@Injectable()
export class TicketsService {
  private tickets: TicketDTO[] = [
    {
      id: 'TK-1001',
      matchTitle: 'MAROC vs FRANCE',
      buyerName: 'Karim Bennani',
      tribune: 'Tribune Nord',
      seatCode: 'NORD-R03-S14',
      price: 150,
      qrHash: 'ETK-2026-X9F4A7B2',
      status: 'VALIDE'
    },
    {
      id: 'TK-1002',
      matchTitle: 'MAROC vs FRANCE',
      buyerName: 'Sara El Amrani',
      tribune: 'VIP Lounge',
      seatCode: 'VIP-R01-S04',
      price: 600,
      qrHash: 'ETK-2026-V88B11C4',
      status: 'VALIDE'
    },
    {
      id: 'TK-1003',
      matchTitle: 'WAC vs RAJA',
      buyerName: 'Youssef Tazi',
      tribune: 'Tribune Est',
      seatCode: 'EST-R08-S22',
      price: 250,
      qrHash: 'ETK-2026-Z77F99A1',
      status: 'SCANNE',
      scannedAt: '20:14:02'
    }
  ];

  findAll() {
    return this.tickets;
  }

  findByQrHash(hash: string) {
    return this.tickets.find(t => t.qrHash === hash);
  }

  issueTicket(matchTitle: string, buyerName: string, tribune: string, seatCode: string, price: number): TicketDTO {
    const randomHash = `ETK-2026-${Math.random().toString(36).substring(2, 10).toUpperCase()}`;
    const newTicket: TicketDTO = {
      id: `TK-${Date.now().toString().slice(-5)}`,
      matchTitle,
      buyerName,
      tribune,
      seatCode,
      price,
      qrHash: randomHash,
      status: 'VALIDE'
    };

    this.tickets.push(newTicket);
    return newTicket;
  }

  validateTicket(hash: string, gateCode: string) {
    const ticket = this.tickets.find(t => t.qrHash === hash);
    if (!ticket) {
      return {
        isAllowed: false,
        statusCode: 'REJECTED_INVALID',
        message: 'Code QR non reconnu dans la base.'
      };
    }

    if (ticket.status === 'SCANNE') {
      return {
        isAllowed: false,
        statusCode: 'REJECTED_DUPLICATE',
        message: `Billet déjà scanné le ${ticket.scannedAt}`
      };
    }

    ticket.status = 'SCANNE';
    ticket.scannedAt = new Date().toLocaleTimeString();

    return {
      isAllowed: true,
      statusCode: 'GRANTED',
      buyerName: ticket.buyerName,
      tribuneName: ticket.tribune,
      seatCode: ticket.seatCode,
      latencyMs: 118,
      message: 'Accès Accordé'
    };
  }
}
