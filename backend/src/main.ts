import { NestFactory } from '@nestjs/core';
import { ValidationPipe } from '@nestjs/common';
import { MicroserviceOptions, Transport } from '@nestjs/microservices';
import { join } from 'path';
import { AppModule } from './app.module';

async function bootstrap() {
  // Create Main NestJS HTTP Application (REST API)
  const app = await NestFactory.create(AppModule);

  // Enable CORS & Global Validation Pipe
  app.enableCors();
  app.setGlobalPrefix('api/v1');
  app.useGlobalPipes(new ValidationPipe({ whitelist: true, transform: true }));

  // Connect gRPC Microservice Transport for Gate Control (< 0.15s)
  app.connectMicroservice<MicroserviceOptions>({
    transport: Transport.GRPC,
    options: {
      package: 'access_control',
      protoPath: join(__dirname, '../proto/access_control.proto'),
      url: '0.0.0.0:50051',
    },
  });

  // Start Microservices and HTTP Server
  await app.startAllMicroservices();
  const port = process.env.PORT || 4000;
  await app.listen(port);

  console.log(`🚀 E-Ticket Pro Production Backend running on REST http://localhost:${port}/api/v1`);
  console.log(`⚡ gRPC High-Speed Gate Control Service listening on 0.0.0.0:50051`);
}
bootstrap();
