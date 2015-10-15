# -*- coding: utf-8 -*-

import sys
import re
import urllib
import base64
import quopri

from cStringIO import StringIO

# ※ email.header だと、subject 等がうまくデコードできない不具合があったため、パッチをあてた版を使用
#    参照： [メール処理でいろいろとはまる - 風柳メモ](http://furyu.hatenablog.com/entry/20140115/1389793767)
#from email.header import decode_header, make_header
from header import decode_header, make_header

from email import message_from_string
from email.utils import getaddresses
from email.generator import Generator

class email_decoder( object ): #{

  #{ // 定数
  _mailaddr_key = { 'from' : True, 'to' : True, 'cc' : True, 'bcc' : True, 'reply-to' : True, 'sender' : True, }
  _one_line_key = { 'subject' : True, 'date' : True, }
  #}
  
  #{ // 変数
  _encoding_for_print = 'utf-8'
  _decode_mime = True
  _decode_char = True
  _message = None
  _message_dict = []
  _body = None
  _html = None
  #}
  
  def log( self, *args ): #{
    if not self.debug:
      return
    for arg in args:
      if isinstance( arg, unicode ):
        arg = arg.encode( self._encoding_for_print, 'replace' )
      print arg
  #}
  
  def logerr( self, *args ): #{
    for arg in args:
      if isinstance( arg, unicode ):
        arg = arg.encode( self._encoding_for_print, 'replace' )
      print >> sys.stderr, arg
  #}
  
  def __init__( self, mime_string_or_file = '', decode_mime = True, decode_char = True, debug = False ): #{
    self.debug = debug
    encoding = sys.getfilesystemencoding()
    if encoding:
      self._encoding_for_print = encoding
    
    self._decode_mime = decode_mime
    self._decode_char = decode_char
    
    if hasattr( mime_string_or_file, 'read' ):
      mime_string = mime_string_or_file.read()
    else:
      mime_string = mime_string_or_file
    
    message = self._message = message_from_string( mime_string )
    
    message_dict = self._message_dict = dict(
      body = [],
      html = [],
      attachments = [],
    )
    mbody = message_dict[ 'body' ]
    mhtml = message_dict[ 'html' ]
    mattach = message_dict[ 'attachments' ]
    _decode = self._decode
    
    is_first_part = True
    next_part_is_multipart_file_content = False
    
    for part in message.walk():
      _ctype = part.get_content_type()
      self.log( '*** Content-Type: %s' % (_ctype) )
      
      part_message_dict = {}
      
      if is_first_part:
        is_first_part = False
        for _key in part.keys():
          _lkey = _key.lower()
          _val_list = [ _decode( _val ) for _val in part.get_all( _key ) ]
          
          part_message_dict.setdefault( _lkey, [] )
          part_message_dict[ _lkey ] += _val_list
          message_dict.setdefault( _lkey, [] )
          message_dict[ _lkey ] += _val_list
          
          self.log( u'%s = %s' % (_lkey, u', '.join(message_dict[ _lkey ])) )
      
      if part.is_multipart():
        self.log( u'----- multipart -----' )
        filename = part.get_filename( failobj = None )
        if not filename:
          continue
        
        next_part_is_multipart_file_content = True
        #attached_message = part.get_payload( i = 0 ).as_string()
        attached_message = self._flatten( part.get_payload( i = 0 ) )
        
        content_transfer_encoding = part_message_dict.get( 'content-transfer-encoding', [ None ] )[ 0 ]
        if content_transfer_encoding == 'base64':
          attached_message = base64.b64decode( attached_message )
        elif content_transfer_encoding == 'quoted-printable':
          attached_message = quopri.decodestring( attached_message )
        
        mattach.append( self._get_attachment_info( part, filename, attached_message ) )
      else:
        self.log( u'----- not multipart (string) -----' )
        if next_part_is_multipart_file_content:
          next_part_is_multipart_file_content = False
          continue
        
        filename = part.get_filename( failobj = None )
        
        payload = part.get_payload( decode = decode_mime )
        
        if decode_char:
          charset = part.get_content_charset( failobj = None )
          if charset:
            try:
              payload = unicode( payload, charset )
            except Exception, s:
              if re.search( '^iso[\-_]?2022[\-_]jp$', charset, re.I ):
                # ISO-2022-JP とあってもこれに含まれない文字を使っているケースがあるため、ISO-2022-JP-2004 とみなして試行
                payload = unicode( payload.replace( '\x1b$B', '\x1b$(Q' ), 'ISO-2022-JP-2004', 'replace' )
        
        if filename:
          payload = part.get_payload( decode = True )
          mattach.append( self._get_attachment_info( part, filename, payload ) )
        else:
          if _ctype == 'text/plain':
            mbody.append( payload )
            self.log( payload )
          elif _ctype == 'text/html':
            mhtml.append( payload )
            self.log( payload )
      
    if not message_dict.has_key( 'sender' ) and message_dict.has_key( 'from' ):
      message_dict[ 'sender' ] = message_dict[ 'from' ]
    
    self._body = u''.join( mbody )
    self._html = u''.join( mhtml )
  #}
  
  def __getattr__( self, name, *opt ): #{
    name = name.lower()
    try:
      attr = super( email_decoder, self ).__getattr__( name )
    except:
      if not self._message_dict.has_key( name ):
        if 0 < len( opt ):
          return opt[ 0 ]
        else:
          raise AttributeError, name
      
      attr = self._message_dict.get( name )
      
      if self._mailaddr_key.get( name ):
        address_list = []
        for ( _realname, _email_address ) in self.list_addresses( name, address_only = False ):
          if _realname:
            address_list.append( u'%s <%s>' % ( _realname, _email_address ) )
          else:
            address_list.append( _email_address )
        attr = u'; '.join( address_list )
      
      if self._one_line_key.get( name ):
        attr = u''.join( attr )
    return attr
  #}
  
  def _decode( self, payload ): #{
    if not self._decode_char:
      return payload
    
    _dclist = decode_header( payload )
    
    try:
      _header = make_header( _dclist )
      _texts = []
      for ( s, charset ) in _header._chunks:
        _texts.append( unicode( s, str( charset ) ).strip() )
      return u''.join( _texts )
    except Exception, s:
      self.logerr( u'_decode():', s )
      _texts = []
      for ( _dc, _code ) in _dclist:
        try:
          if _code:
            _dc = unicode( _dc, encoding = _code )
          else:
            _dc = unicode( _dc )
        except Exception, s:
          self.logerr( u'email_decoder()._decode()', [ _dc ] )
          try:
            _dc = urllib.unquote( _dc ).decode( 'utf-8' )
          except Exception, s:
            self.logerr( u'email_decoder()._decode()', [ _dc ] )
            _dc = unicode( _dc, errors = 'ignore' )
        _texts.append( _dc.strip() )
      
      return u''.join( _texts )
  #}
  
  def _flatten( self, message ): #{
    # [参考] http://docs.python.jp/2/library/email.message.html#email.message.Message.as_string
    fp = StringIO()
    gen = Generator( fp, mangle_from_ = False, maxheaderlen = 60 )
    gen.flatten( message )
    return fp.getvalue()
  #}
  
  def _get_attachment_info( self, part, encoded_filename, payload ): #{
    filename = self._decode( encoded_filename )
    filename_with_extension = encoded_filename
    ctype = part.get_content_type()
    ctype_main = part.get_content_maintype()
    ctype_sub = part.get_content_subtype()
    if not re.search( u'\.[^.]*$', filename_with_extension ) and ctype_sub:
      filename_with_extension = filename_with_extension + '.' + ctype_sub
    
    self.log( u'filename: %s' % (filename) )
    
    return dict(
      filename = filename,
      filename_with_extension = filename_with_extension,
      ctype = ctype,
      ctype_main =  ctype_main,
      ctype_sub = ctype_sub,
      payload = payload,
    )
  #}
  
  def is_decoded_mime( self ): #{
    return self._decode_mime
  #}
  
  def is_decoded_char( self ): #{
    return self._decode_char
  #}
  
  def get_original_message( self ): #{
    return self._message
  #}
  
  def get_bodies( self, content_type = 'text' ): #{
    if content_type == 'text/plain':
      return self._body
    elif content_type == 'text/html':
      return self._html
    else:
      return { 'text/plain' : self._body, 'text/html' : self._html }
  #}
  
  def get_body_plain( self ): #{
    #return self.get_bodies( 'text/plain' )
    return self._body
  #}
  
  def get_body_html( self ): #{
    #return self.get_bodies( 'text/html' )
    return self._html
  #}
  
  def bodies( self, content_type = 'text' ): #{
    if content_type == 'text/plain':
      return ( 'text/plain', self._body )
    elif content_type == 'text/html':
      return ( 'text/html', self._html )
    else:
      return [ ('text/plain', self._body), ('text/html', self._html) ]
  #}
  
  def list_addresses( self, name, address_only = True ): #{
    name = name.lower()
    if not self._mailaddr_key.get( name ) or not self._message_dict.has_key( name ):
      return []
    all_addresses = getaddresses( self._message_dict.get( name ) )
    if address_only:
      return [ _email for ( _realname, _email ) in all_addresses ]
    else:
      return all_addresses
  #}
  
  def list_attribute_names( self ): #{
    return self._message_dict.keys()
  #}

#} // end of class email_decoder()


# ■ end of file
