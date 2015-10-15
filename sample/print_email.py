# -*- coding: utf-8 -*-

from __future__ import print_function

import os
import sys
import re

from email.utils import parsedate_tz

include_path = os.path.join( os.path.dirname( os.path.abspath(__file__) ), '..' )
sys.path.append( include_path )

from email_decoder import email_decoder


system_encoding = sys.getfilesystemencoding()

def prn( *args ): #{
  for arg in args:
    if isinstance( arg, unicode ):
      arg = arg.encode( system_encoding, 'replace' )
    print( arg, end = '' )
  print( '' )
#}


def print_email( mail_text_or_file ): #{
  mail = email_decoder( mail_text_or_file )
  
  # -- 属性一覧
  attribute_names = mail.list_attribute_names()
  prn( u'属性一覧: ', '\n', attribute_names, '\n' )
  
  # -- From
  prn( 'From: ', mail.sender ) # 'from' は Python の予約語のため、mail.from はエラーになる
  
  # -- To
  prn( 'To: ', mail.to )
  
  # -- Cc
  if hasattr( mail, 'cc' ):
    prn( 'Cc: ', mail.cc )
    
    ''' # メールアドレスだけを配列で取り出す例
    for mail_address in mail.list_addresses( 'cc' ):
      prn( 'Cc: ', mail_address )
    '''
    ''' # 名前とメールアドレスペアの配列を取り出す例
    for ( name, mail_address ) in mail.list_addresses( 'cc', address_only = False ):
      prn( 'Cc: ', u'%s <%s>' % ( name, mail_address ) )
    '''
  
  # -- Bcc
  if hasattr( mail, 'bcc' ):
    prn( 'Bcc: ', mail.bcc )
  
  # -- Reply-To
  if hasattr( mail, 'reply-to' ):
    prn( 'Reply-To: ', getattr( mail, 'reply-to' ) )
  
  prn( '' )
  
  # -- 日付・時刻
  if hasattr( mail, 'date' ):
    ( tm_year, tm_mon, tm_mday, tm_hour, tm_min, tm_sec, tm_wday, tm_yday, tm_isdst, tm_tz ) = parsedate_tz( mail.date )
    date_time = '%04d-%02d-%02d %02d:%02d:%02d' % ( tm_year, tm_mon, tm_mday, tm_hour, tm_min, tm_sec )
    prn( u'日付: ', date_time, '\n' )
  
  # -- 表題
  subject = mail.subject
  prn( u'表題: ', subject, '\n' )
  
  # -- 本文(平文)
  body = mail.get_body_plain()
  if body:
    prn( u'本文(平文): ', '\n', body, '\n' )
  
  # -- 本文(HTML)
  body_html = mail.get_body_html()
  if body_html:
    prn( u'本文(HTML): ', '\n', body_html, '\n')
  
  # -- 添付ファイル書き出し
  attachments = mail.attachments
  for attachment in attachments:
    filename = attachment[ 'filename' ]
    prn( u'添付ファイル: ', filename )
    encoded_filename = filename.encode( system_encoding, 'replace' )
    if os.path.exists( encoded_filename ): # 上書き防止
      prn( u'→すでに "%s" が存在します' % ( filename ) )
      continue
    fp = open( encoded_filename, 'wb' )
    fp.write( attachment[ 'payload' ] )
    fp.close()
#}


def print_email_file( mail_filename ): #{
  with open( mail_filename, 'rb' ) as file:
    prn( u'■ E-Mailファイル: ', mail_filename )
    prn( '-' * 80 )
    print_email( file )
    prn( '-' * 80, '\n' )
#}

if __name__ == '__main__': #{
  if len( sys.argv ) < 2:
    print_email( sys.stdin.read() )
  else:
    for file_or_directory in sys.argv[1:]:
      if os.path.isdir( file_or_directory ):
        for filename in os.listdir( file_or_directory ):
          if not re.search( '\.eml$', filename ):
            continue
          print_email_file( os.path.join( file_or_directory, filename ) )
      elif os.path.isfile( file_or_directory ):
        print_email_file( file_or_directory )
      else:
        prn( file_or_directory, u' → 見つかりません' )
#}

# ■ end of file
