"""Generate the raster icons referenced by manifest.json.

The PNG outputs are intentionally gitignored because the PR system accepts
text files only. Run this script before loading the unpacked extension.
"""
import struct,zlib,math
S=4
FONT={'P':['11110','10001','11110','10000','10000','10000','10000'],'D':['11110','10001','10001','10001','10001','10001','11110'],'F':['11111','10000','11110','10000','10000','10000','10000']}
def make(n):
 w=h=n*S; p=[[(0,0,0,0) for _ in range(w)] for _ in range(h)]
 def blend(x,y,c):
  if 0<=x<w and 0<=y<h:p[y][x]=c
 def rr(x0,y0,x1,y1,r,c):
  for y in range(y0,y1):
   for x in range(x0,x1):
    cx=min(max(x,x0+r),x1-r-1); cy=min(max(y,y0+r),y1-r-1)
    if (x-cx)**2+(y-cy)**2<=r*r: blend(x,y,c)
 def poly(pts,c):
  miny,maxy=int(min(y for x,y in pts)),int(max(y for x,y in pts))
  for y in range(miny,maxy+1):
   xs=[]
   for (x1,y1),(x2,y2) in zip(pts,pts[1:]+pts[:1]):
    if (y1<=y<y2) or (y2<=y<y1): xs.append(x1+(y-y1)*(x2-x1)/(y2-y1))
   xs.sort()
   for a,b in zip(xs[::2],xs[1::2]):
    for x in range(math.ceil(a),math.floor(b)+1): blend(x,y,c)
 def ellipse(cx,cy,rx,ry,c):
  for y in range(int(cy-ry),int(cy+ry)+1):
   for x in range(int(cx-rx),int(cx+rx)+1):
    if ((x-cx)/rx)**2+((y-cy)/ry)**2<=1:blend(x,y,c)
 u=lambda v:round(v*n*S/128)
 # distinct Outlook-blue inspired tile and envelope
 rr(u(8),u(12),u(116),u(116),u(23),(0,120,212,255))
 poly([(u(13),u(51)),(u(62),u(83)),(u(111),u(51)),(u(111),u(104)),(u(13),u(104))],(16,96,176,255))
 # white O ring
 ellipse(u(42),u(55),u(24),u(25),(255,255,255,255)); ellipse(u(42),u(55),u(12),u(13),(0,120,212,255))
 # document with folded corner
 rr(u(64),u(25),u(109),u(91),u(5),(235,248,255,255))
 poly([(u(92),u(25)),(u(109),u(42)),(u(92),u(42))],(80,190,235,255))
 # PDF badge
 rr(u(55),u(76),u(116),u(105),u(7),(0,73,135,255))
 if n>=32:
  scale=max(1,u(2)); gap=u(2); total=3*5*scale+2*gap; x0=u(85)-total//2; y0=u(83)
  x=x0
  for ch in 'PDF':
   for yy,row in enumerate(FONT[ch]):
    for xx,v in enumerate(row):
     if v=='1':
      for dy in range(scale):
       for dx in range(scale): blend(x+xx*scale+dx,y0+yy*scale+dy,(255,255,255,255))
   x+=5*scale+gap
 # box downsample
 out=[]
 for y in range(n):
  row=[]
  for x in range(n):
   vals=[p[y*S+dy][x*S+dx] for dy in range(S) for dx in range(S)]
   row.append(tuple(sum(v[i] for v in vals)//len(vals) for i in range(4)))
  out.append(row)
 raw=b''.join(b'\x00'+b''.join(bytes(px) for px in row) for row in out)
 def chunk(t,d):return struct.pack('>I',len(d))+t+d+struct.pack('>I',zlib.crc32(t+d)&0xffffffff)
 png=b'\x89PNG\r\n\x1a\n'+chunk(b'IHDR',struct.pack('>IIBBBBB',n,n,8,6,0,0,0))+chunk(b'IDAT',zlib.compress(raw,9))+chunk(b'IEND',b'')
 import os
 os.makedirs('icons', exist_ok=True)
 open(f'icons/icon-{n}.png','wb').write(png)
for n in (16,32,48,128):make(n)
