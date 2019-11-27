.PHONY: index spec bom gspec gbom

default: all

all: index spec bom gspec gbom

index:
	cd index && \
	cd Scripts/python/doc/ && \
	asciidoctor help.adoc && \
	cd - && \
	zip -FS -r ~/.config/libreoffice/4/user/template/Перечень\ элементов.ott * -x Scripts/python/doc/help.adoc

spec:
	cd spec && \
	cd Scripts/python/doc/ && \
	asciidoctor help.adoc && \
	cd - && \
	zip -FS -r ~/.config/libreoffice/4/user/template/Спецификация.ott * -x Scripts/python/doc/help.adoc

bom:
	cd bom && \
	cd Scripts/python/doc/ && \
	asciidoctor help.adoc && \
	cd - && \
	zip -FS -r ~/.config/libreoffice/4/user/template/Ведомость\ покупных\ изделий.ott * -x Scripts/python/doc/help.adoc

gspec:
	cd gspec && \
	cd Scripts/python/doc/ && \
	asciidoctor help.adoc && \
	cd - && \
	zip -FS -r ~/.config/libreoffice/4/user/template/Групповая\ спецификация.ott * -x Scripts/python/doc/help.adoc

gbom:
	cd gbom && \
	cd Scripts/python/doc/ && \
	asciidoctor help.adoc && \
	cd - && \
	zip -FS -r ~/.config/libreoffice/4/user/template/Групповая\ ведомость\ покупных\ изделий.ott * -x Scripts/python/doc/help.adoc
