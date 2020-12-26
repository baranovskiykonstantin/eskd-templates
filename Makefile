OUT := ~/.config/libreoffice/4/user/template
VERSION := 1.4

.PHONY: index spec bom gspec gbom manual mexanic archive

default: all

all: index spec bom gspec gbom manual mexanic

define build_ott
	cd $(1) && \
	cd Scripts/python/doc/ && \
	asciidoctor help.adoc && \
	cd - && \
	zip -FS -r $(OUT)/$(2).ott * -x Scripts/python/doc/help.adoc
endef

index:
	$(call build_ott,index,Перечень\ элементов)

spec:
	$(call build_ott,spec,Спецификация)

bom:
	$(call build_ott,bom,Ведомость\ покупных\ изделий)

gspec:
	$(call build_ott,gspec,Групповая\ спецификация)

gbom:
	$(call build_ott,gbom,Групповая\ ведомость\ покупных\ изделий)

manual:
	$(call build_ott,manual,Пояснительная\ записка)

mexanic:
	$(call build_ott,mexanic,Ведомость\ покупных\ изделий\ \(Mexanic\))

archive:
	cd $(OUT) && \
	7z a -- /tmp/eskd-templates_v$(VERSION).7z \
	Перечень\ элементов.ott \
	Спецификация.ott \
	Ведомость\ покупных\ изделий.ott \
	Групповая\ спецификация.ott \
	Групповая\ ведомость\ покупных\ изделий.ott \
	Пояснительная\ записка.ott \
	Ведомость\ покупных\ изделий\ \(Mexanic\).ott
