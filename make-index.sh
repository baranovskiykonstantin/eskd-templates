cd index
cd Scripts/python/index/doc/
asciidoctor help.adoc
cd -
zip -r ~/.config/libreoffice/4/user/template/Перечень\ элементов.ott * -x Scripts/python/index/doc/help.adoc
