"""
printer_relay.py - TCP print relay for Windows spool jobs.

Listens on a local TCP port, accepts one raw print job per connection,
converts the incoming bytes to the printer's native payload, and sends
the result to the saved BLE printer config from bt_scan.py.
"""

import argparse
import asyncio
import itertools
import os
import platform
import sys

import bt_print
from bt_shared import BLEAK_AVAILABLE, CONFIG_FILE, header, info, fail, ok, load_config


class RelayServer:
    def __init__(
        self,
        host: str,
        port: int,
        idle_timeout: float,
        max_job_bytes: int,
        cfg: dict,
    ) -> None:
        self.host = host
        self.port = port
        self.idle_timeout = idle_timeout
        self.max_job_bytes = max_job_bytes
        self.job_counter = itertools.count(1)
        self.send_lock = asyncio.Lock()
        self.cfg = cfg
        self.use_cat = bt_print._is_cat_printer(cfg.get("write_uuid", ""))
        self._config_mtime = self._read_config_mtime()

    async def run(self) -> None:
        server = await asyncio.start_server(self._handle_client, self.host, self.port)
        sockets = ", ".join(self._format_socket(sock.getsockname()) for sock in server.sockets or [])
        ok(f"Relay listening on {sockets or f'{self.host}:{self.port}'}")
        info("Waiting for Windows print jobs...")
        async with server:
            await server.serve_forever()

    async def _handle_client(
        self, reader: asyncio.StreamReader, writer: asyncio.StreamWriter
    ) -> None:
        job_num = next(self.job_counter)
        peer = writer.get_extra_info("peername")
        peer_label = self._format_socket(peer) if peer else "unknown"
        info(f"Connection accepted for job #{job_num} from {peer_label}")

        try:
            raw_job = await self._read_job(reader, job_num)
            if not raw_job:
                fail(f"Job #{job_num}: no bytes received")
                writer.write(b"ERROR empty job\n")
                await writer.drain()
                return

            async with self.send_lock:
                success = await self._process_job(job_num, raw_job)

            writer.write(b"OK\n" if success else b"ERROR relay failed\n")
            await writer.drain()
        except Exception as exc:
            fail(f"Job #{job_num}: relay exception: {exc}")
            try:
                writer.write(f"ERROR {exc}\n".encode("utf-8", errors="replace"))
                await writer.drain()
            except Exception:
                pass
        finally:
            writer.close()
            try:
                await writer.wait_closed()
            except (ConnectionError, OSError):
                pass

    async def _read_job(self, reader: asyncio.StreamReader, job_num: int) -> bytes:
        chunks: list[bytes] = []
        total = 0
        idle_timeouts = 0

        while True:
            try:
                chunk = await asyncio.wait_for(reader.read(65536), timeout=self.idle_timeout)
            except asyncio.TimeoutError:
                if chunks:
                    info(
                        f"Job #{job_num}: idle timeout reached after "
                        f"{self.idle_timeout:.1f}s, closing current job"
                    )
                    break
                idle_timeouts += 1
                if idle_timeouts >= 2:
                    info(
                        f"Job #{job_num}: no data received after "
                        f"{self.idle_timeout * idle_timeouts:.1f}s, closing connection"
                    )
                    break
                continue

            if not chunk:
                break

            chunks.append(chunk)
            total += len(chunk)
            idle_timeouts = 0
            if total > self.max_job_bytes:
                raise RuntimeError(
                    f"job exceeds max size ({total} bytes > {self.max_job_bytes} bytes)"
                )

        return b"".join(chunks)

    async def _process_job(self, job_num: int, raw_job: bytes) -> bool:
        cfg = self._get_config()
        if not cfg:
            fail(f"No printer config at {CONFIG_FILE}")
            info("Run bt_scan.py --save first.")
            return False

        fmt = bt_print.detect_format(raw_job)

        info(f"Job #{job_num}: received {len(raw_job)} bytes")
        info(f"Job #{job_num}: detected format = {fmt}")
        info(
            f"Job #{job_num}: target protocol = "
            f"{'BLE_CAT' if self.use_cat else 'ESC/POS'}"
        )

        if self.use_cat:
            payload = bt_print._to_cat_payload(raw_job, fmt, job_num)
        else:
            payload = bt_print.convert_to_escpos(raw_job, fmt)

        if not payload:
            fail(f"Job #{job_num}: conversion produced empty payload")
            return False

        info(f"Job #{job_num}: converted payload size = {len(payload)} bytes")
        success = await bt_print.send_direct_ble(payload, cfg)
        if success:
            ok(f"Job #{job_num}: delivered to BLE printer")
        else:
            fail(f"Job #{job_num}: BLE delivery failed")
        return success

    @staticmethod
    def _format_socket(sockname: object) -> str:
        if isinstance(sockname, tuple):
            if len(sockname) >= 2:
                return f"{sockname[0]}:{sockname[1]}"
        return str(sockname)

    @staticmethod
    def _read_config_mtime() -> float | None:
        try:
            return os.path.getmtime(CONFIG_FILE)
        except OSError:
            return None

    def _get_config(self) -> dict | None:
        current_mtime = self._read_config_mtime()
        if current_mtime != self._config_mtime:
            cfg = load_config()
            if not cfg:
                self.cfg = None
                self.use_cat = False
                self._config_mtime = current_mtime
                return None
            self.cfg = cfg
            self.use_cat = bt_print._is_cat_printer(cfg.get("write_uuid", ""))
            self._config_mtime = current_mtime
            info("Printer config reloaded from disk")
        return self.cfg


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="printer_relay.py - TCP relay for Windows print jobs",
    )
    parser.add_argument(
        "--host",
        default="127.0.0.1",
        help="Host/IP to bind. Default: 127.0.0.1",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=9100,
        help="TCP port to bind. Default: 9100",
    )
    parser.add_argument(
        "--idle-timeout",
        type=float,
        default=1.5,
        help="Seconds of socket inactivity that marks end-of-job. Default: 1.5",
    )
    parser.add_argument(
        "--max-job-mb",
        type=float,
        default=32.0,
        help="Reject jobs larger than this many MiB. Default: 32",
    )
    return parser.parse_args(argv)


async def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)

    header("printer_relay.py - TCP print relay")
    print(f"  Platform : {platform.system()} {platform.release()}")
    print(f"  Python   : {sys.version.split()[0]}")
    print(f"  Config   : {CONFIG_FILE}")
    print(f"  Bind     : {args.host}:{args.port}")
    print(f"  Timeout  : {args.idle_timeout:.1f}s")

    if not BLEAK_AVAILABLE:
        fail("'bleak' not installed. Run: pip install bleak")
        return 1

    if args.port < 1 or args.port > 65535:
        fail("Port must be between 1 and 65535.")
        return 1
    if args.idle_timeout <= 0:
        fail("Idle timeout must be greater than 0.")
        return 1
    if args.max_job_mb <= 0:
        fail("Max job size must be greater than 0.")
        return 1

    cfg = load_config()
    if not cfg:
        fail(f"No printer config at {CONFIG_FILE}")
        info("Run bt_scan.py --save first.")
        return 1

    max_job_bytes = max(1, int(args.max_job_mb * 1024 * 1024))
    relay = RelayServer(args.host, args.port, args.idle_timeout, max_job_bytes, cfg)
    await relay.run()
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(asyncio.run(main()))
    except KeyboardInterrupt:
        print()
        info("Relay stopped by user.")
