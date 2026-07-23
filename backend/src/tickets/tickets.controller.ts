import { Controller, Get, Post, Body, Param } from '@nestjs/common';
import { TicketsService } from './tickets.service';

@Controller('tickets')
export class TicketsController {
  constructor(private readonly ticketsService: TicketsService) {}

  @Get()
  getAllTickets() {
    return this.ticketsService.findAll();
  }

  @Get(':hash')
  getTicketByHash(@Param('hash') hash: string) {
    return this.ticketsService.findByQrHash(hash);
  }

  @Post('issue')
  issueTicket(@Body() body: { matchTitle: string; buyerName: string; tribune: string; seatCode: string; price: number }) {
    return this.ticketsService.issueTicket(
      body.matchTitle,
      body.buyerName,
      body.tribune,
      body.seatCode,
      body.price
    );
  }

  @Post('validate')
  validateTicket(@Body() body: { qrHash: string; gateCode: string }) {
    return this.ticketsService.validateTicket(body.qrHash, body.gateCode);
  }
}
