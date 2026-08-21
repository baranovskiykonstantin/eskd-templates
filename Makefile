OUT := ~/.config/libreoffice/4/user/template
VERSION := $(file < version)

# Проверка наличия asciidoctor и zip (обязательные для сборки)
REQUIRED := asciidoctor zip
$(foreach cmd,$(REQUIRED), \
    $(if $(shell command -v $(cmd) 2>/dev/null),, \
        $(error "Не найдена утилита '$(cmd)'. Установите её (см. README)") \
    ) \
)

.PHONY: index spec bom gspec gbom manual mexanic archive

default: all

all: index spec bom gspec gbom manual mexanic

define build_ott
	cd $(1) && \
	cd Scripts/python/doc/ && \
	asciidoctor help.adoc && \
	cd - && \
	mkdir -p $(OUT) && \
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
	@command -v 7z >/dev/null 2>&1 || { echo "Ошибка: 7z не найден. Установите p7zip (Ubuntu: sudo apt install p7zip-full)"; exit 1; }
	cd $(OUT) && \
	7z a -- /tmp/eskd-templates_v$(VERSION).7z \
	Перечень\ элементов.ott \
	Спецификация.ott \
	Ведомость\ покупных\ изделий.ott \
	Групповая\ спецификация.ott \
	Групповая\ ведомость\ покупных\ изделий.ott \
	Пояснительная\ записка.ott \
	Ведомость\ покупных\ изделий\ \(Mexanic\).ott && \
	cd - && \
	7z a -- /tmp/example_v$(VERSION).7z example/
