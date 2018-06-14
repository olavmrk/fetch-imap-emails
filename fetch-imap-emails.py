#!/usr/bin/env python3
import argparse
import contextlib
import csv
import email
import email.policy
import imaplib
import sys
import urllib.parse


class ImapUrl:
    def __init__(self, url):
        parsed = urllib.parse.urlparse(url)

        if parsed.scheme not in ('imap', 'imaps'):
            raise ValueError('Unknown IMAP scheme: {scheme}'.format(scheme=parsed.scheme))
        self.ssl = parsed.scheme == 'imaps'

        if not parsed.hostname:
            raise ValueError('Missing host in IMAP url')
        self.hostname = parsed.hostname

        if not parsed.hostname:
            raise ValueError('Missing host in IMAP url')
        self.hostname = parsed.hostname

        if parsed.port is not None:
            self.port = parsed.port
        else:
            self.port = imaplib.IMAP4_SSL_PORT if self.ssl else imaplib.IMAP4_PORT

        if not parsed.path:
            raise ValueError('Missing folder name')
        self.mailbox = parsed.path[1:]  # Skip initial '/'

        self.username = parsed.username
        self.password = parsed.password

    def __repr__(self):
        return 'ImapUrl(ssl={ssl!r}, hostname={hostname!r}, port={port!r}, mailbox={mailbox!r}, username={username!r}, password={password!r})'.format(
            ssl=self.ssl,
            hostname=self.hostname,
            port=self.port,
            mailbox=self.mailbox,
            username=self.username,
            password=self.password,
        )


class FieldSpec:
    def __init__(self, fieldspec):
        if ':' in fieldspec:
            name, decoder = fieldspec.split(':', 1)
        else:
            name = fieldspec
            decoder = 'noop'
        self.name = name

        decoder_method = '_decoder_' + decoder
        self.decode = getattr(FieldSpec, decoder_method, None)
        if not self.decode:
            raise ValueError('Invalid decoder: {decoder}'.format(decoder=decoder))

    @staticmethod
    def parse_arg(fields):
        return [ FieldSpec(fieldspec) for fieldspec in fields.split(',') ]

    @staticmethod
    def _decoder_noop(value):
        return value

    @staticmethod
    def _decoder_isotime(value):
        return email.utils.parsedate_to_datetime(value).isoformat()

    @staticmethod
    def _decoder_address(value):
        return email.utils.parseaddr(value)[1].lower()


def parse_args():
    parser = argparse.ArgumentParser(description='Fetch email fields as csv')
    parser.add_argument('--fields', type=FieldSpec.parse_arg, help='Comma-separated list of fields to fetch', default='from,to,date,subject')
    parser.add_argument('target', action='store', type=ImapUrl, help='IMAP mailbox, as an URL (imaps://username:password@hostname/folder)')
    return parser.parse_args()


@contextlib.contextmanager
def imap_connection(parsed_url):
    if parsed_url.ssl:
        connection = imaplib.IMAP4_SSL(host=parsed_url.hostname, port=parsed_url.port)
    else:
        connection = imaplib.IMAP4(host=parsed_url.hostname, port=parsed_url.port)
    with connection:
        if parsed_url.username or parsed_url.password:
            connection.login(parsed_url.username, parsed_url.password)
        code, data = connection.select(mailbox=parsed_url.mailbox.encode('utf-8'), readonly=True)
        if code != 'OK':
            raise Exception('Selecting mailbox {mailbox} failed: {code} {data!r}'.format(mailbox=args.target.mailbox, code=code, data=data))
        yield connection


def fetch_messages(connection, fields):
    code, data = connection.search('UTF-8', 'ALL')
    if code != 'OK':
        raise Exception('Search failed: {code} {data!r}'.format(code=code, data=data))
    message_ids = data[0].split()
    for message_id in message_ids:
        code, data = connection.fetch(message_id, 'BODY[HEADER.FIELDS ({fields})]'.format(fields=' '.join(field.name.upper() for field in fields)))
        if code != 'OK':
            raise Exception('fetch failed: {code} {data!r}'.format(code=code, data=data))
        message_raw = data[0][1]
        message = email.message_from_bytes(message_raw, policy=email.policy.default)
        yield message


def main():
    args = parse_args()

    with imap_connection(args.target) as connection:
        writer = csv.writer(sys.stdout)
        writer.writerow([field.name for field in args.fields])
        for message in fetch_messages(connection, args.fields):
            writer.writerow([field.decode(message.get(field.name, '')) for field in args.fields])


if __name__ == '__main__':
    main()
