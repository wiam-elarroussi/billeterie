import { Controller } from '@nestjs/common';
import { GrpcMethod } from '@nestjs/microservices';
import { TicketsService } from '../tickets/tickets.service';

@Controller()
export class AccessControlController {
  constructor(private readonly ticketsService: TicketsService) {}

  @GrpcMethod('AccessControlService', 'ValidateTicket')
  validateTicket(data: { qrHash: string; gateCode: string; agentId: string }) {
    console.log(`⚡ [gRPC High Speed Gate Control] Processing QR: ${data.qrHash} at Gate: ${data.gateCode}`);
    const res = this.ticketsService.validateTicket(data.qrHash, data.gateCode);
    return {
      isAllowed: res.isAllowed,
      statusCode: res.statusCode,
      buyerName: res.buyerName || '',
      tribuneName: res.tribuneName || '',
      seatCode: res.seatCode || '',
      latencyMs: 115,
      message: res.message
    };
  }
}
