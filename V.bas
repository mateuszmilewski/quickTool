Attribute VB_Name = "V"
'The MIT License (MIT)
'
'Copyright (c) 2017 FORREST
' Mateusz Milewski mateusz.milewski@opel.com aka FORREST
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 2017-11-13
' v0.1 init on this project
' 3 cfg sheets: input, register, plt-list
' OOP schema ICorail -> Corail Blue & Orange - a plan
' also plan to have app.run (kind of multi-thread app)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 2017-11-14
' v0.2 next steps with new classes:
' parser
' rawdata
' shellhandler
' eventhandler connected with corail handler
' sets of corails
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 2017-11-16
' dopisanie implemenacji odpowiedzialnej za frame:
' Set .frame = .doc.frames(QT.G_MAIN_FRAME_ID)
' okazalo sie ze orange corail jest strona w stronie - musialem to jakos obejsc...
'
' new class: DropperHandler
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 2017-11-20
'v0.4
'duzo zmian
'lacznie z pierwszym udanym polaczeniem z danymi na zywym systemie
'jest to pierwsza podwersja pisana bezposrednio na francuskim sprzecie
'testy natychmiastowe bez koniecznosci przeklikiwania sie pomiedzy mailami
' poprawiony parser
' ujednolicone dzialania pomiedzy corailami blue and orange
' schema:
'CorailHelper -> CorailRunner -> ICorail jako interfejs - orange oraz blue korzystaja z tych samych metod

' Orange, Blue, Manual Corail implements ICorail

''w manual Corail wszystkie metody wlasciwie wygldaja tak samo jak w interfejsie - spowodowane jest to glownie brakiem danych pobiernaych
' wiec generalnie jest pusto i cicho - jedyna zmiana to zaprzestanie wyrzucania bledow krytycznych jesli pod koniec logiki dane wciaz
' sa nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 2017-11-21
'v0.5
' nowe funkcje:
' 1 open plants
' 2 close all corails and maestros
' 3 after open plants ie is not visible
' 4 initial layout for the tool
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 2017-11-22
'v0.6
' waiting for IE not working need ta adjust more directly with content of corail site
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' 2018-03-06
' v0.7
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' adjust for safe mode in IE\
' removal of some logic inside layout changes
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



' 2018-03-29
' v0.8
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' add export this project module for githib repository...
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

