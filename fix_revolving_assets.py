from __future__ import annotations

import copy
import zipfile
from pathlib import Path

from lxml import etree
from PIL import Image, ImageDraw, ImageFont


ROOT = Path(r"d:\Data\Code\AutomationBullShit")
DRAWIO_PATH = ROOT / "RevolvingDoor_CyberPhysical.drawio"
DOCX_PATH = ROOT / "RevolvingDoor_Sample.docx"
PNG_PATH = ROOT / "RevolvingDoor_CyberPhysical.png"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"w": W_NS, "r": R_NS, "wp": WP_NS, "a": A_NS, "pic": PIC_NS}


def w_tag(name: str) -> str:
    return f"{{{W_NS}}}{name}"


def load_font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    candidates = []
    if bold:
        candidates.extend(
            [
                r"C:\Windows\Fonts\arialbd.ttf",
                r"C:\Windows\Fonts\calibrib.ttf",
                r"C:\Windows\Fonts\tahomabd.ttf",
            ]
        )
    else:
        candidates.extend(
            [
                r"C:\Windows\Fonts\arial.ttf",
                r"C:\Windows\Fonts\calibri.ttf",
                r"C:\Windows\Fonts\tahoma.ttf",
            ]
        )

    for candidate in candidates:
        path = Path(candidate)
        if path.exists():
            return ImageFont.truetype(str(path), size=size)
    return ImageFont.load_default()


def fix_drawio() -> None:
    drawio_xml = """<mxfile host="app.diagrams.net" modified="2026-03-10T11:05:00.000Z" agent="Codex" version="27.0.5">
  <diagram name="CyberPhysical Diagram" id="revolving-door-cpd">
    <mxGraphModel dx="1600" dy="1000" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="1" pageScale="1" pageWidth="1400" pageHeight="1000" math="0" shadow="0">
      <root>
        <mxCell id="0"/>
        <mxCell id="1" parent="0"/>

        <mxCell id="env_cloud" value="" style="shape=cloud;whiteSpace=wrap;html=1;fillColor=#ffffff;strokeColor=#000000;strokeWidth=1.5;" vertex="1" parent="1">
          <mxGeometry x="70" y="20" width="1260" height="330" as="geometry"/>
        </mxCell>
        <mxCell id="env_label" value="&lt;b&gt;Environment&lt;/b&gt;" style="text;html=1;align=center;verticalAlign=middle;resizable=0;autosize=1;fontSize=18;fontStyle=1;" vertex="1" parent="1">
          <mxGeometry x="595" y="30" width="210" height="30" as="geometry"/>
        </mxCell>
        <mxCell id="users" value="Users" style="shape=umlActor;html=1;whiteSpace=wrap;strokeColor=#000000;fillColor=#ffffff;fontSize=13;verticalLabelPosition=bottom;verticalAlign=top;align=center;" vertex="1" parent="1">
          <mxGeometry x="95" y="105" width="60" height="95" as="geometry"/>
        </mxCell>

        <mxCell id="actions_box" value="&lt;b&gt;Actions&lt;/b&gt;&lt;hr&gt;ApproachSideA, ApproachSideB&lt;br&gt;LeaveSideA, LeaveSideB&lt;br&gt;PressPartition, ReleasePartition&lt;br&gt;PressForwardButton, ReleaseForwardButton&lt;br&gt;PressBackwardButton, ReleaseBackwardButton" style="rounded=0;whiteSpace=wrap;html=1;fillColor=#ffffff;strokeColor=#000000;strokeWidth=1.4;fontSize=13;align=center;verticalAlign=middle;spacing=8;" vertex="1" parent="1">
          <mxGeometry x="180" y="105" width="420" height="190" as="geometry"/>
        </mxCell>
        <mxCell id="effects_box" value="&lt;b&gt;Effects&lt;/b&gt;&lt;hr&gt;Passage availability&lt;br&gt;Visible door movement in selected direction&lt;br&gt;Temporary stop under pressure" style="rounded=0;whiteSpace=wrap;html=1;fillColor=#ffffff;strokeColor=#000000;strokeWidth=1.4;fontSize=13;align=center;verticalAlign=middle;spacing=8;" vertex="1" parent="1">
          <mxGeometry x="635" y="130" width="230" height="140" as="geometry"/>
        </mxCell>
        <mxCell id="sensation_box" value="&lt;b&gt;Sensation&lt;/b&gt;&lt;hr&gt;User sees traffic lights&lt;br&gt;and perceives forward / backward rotation&lt;br&gt;or temporary stop" style="rounded=0;whiteSpace=wrap;html=1;fillColor=#ffffff;strokeColor=#000000;strokeWidth=1.4;fontSize=13;align=center;verticalAlign=middle;spacing=8;" vertex="1" parent="1">
          <mxGeometry x="930" y="120" width="320" height="160" as="geometry"/>
        </mxCell>

        <mxCell id="plant_rect" value="" style="rounded=0;whiteSpace=wrap;html=1;fillColor=#ffffff;strokeColor=#000000;strokeWidth=1.6;" vertex="1" parent="1">
          <mxGeometry x="90" y="390" width="1210" height="300" as="geometry"/>
        </mxCell>
        <mxCell id="plant_label" value="&lt;b&gt;Plant&lt;/b&gt;" style="text;html=1;align=center;verticalAlign=middle;resizable=0;autosize=1;fontSize=18;fontStyle=1;" vertex="1" parent="1">
          <mxGeometry x="655" y="400" width="90" height="30" as="geometry"/>
        </mxCell>

        <mxCell id="ctrl_box" value="&lt;b&gt;Controls&lt;/b&gt;&lt;hr&gt;presenceSideA, presenceSideB &lt;b&gt;OUT&lt;/b&gt;&lt;br&gt;partitionPressure &lt;b&gt;OUT&lt;/b&gt;&lt;hr&gt;rotateForwardButton, rotateBackwardButton &lt;b&gt;OUT&lt;/b&gt;" style="rounded=0;whiteSpace=wrap;html=1;fillColor=#ffffff;strokeColor=#000000;strokeWidth=1.4;fontSize=13;align=center;verticalAlign=middle;spacing=8;" vertex="1" parent="1">
          <mxGeometry x="110" y="435" width="245" height="200" as="geometry"/>
        </mxCell>
        <mxCell id="sens_box" value="&lt;b&gt;Sensors&lt;/b&gt;&lt;hr&gt;Plant &lt;b&gt;FROM&lt;/b&gt; angle / motion sensors&lt;hr&gt;doorAngle, isRotating &lt;b&gt;OUT&lt;/b&gt;" style="rounded=0;whiteSpace=wrap;html=1;fillColor=#ffffff;strokeColor=#000000;strokeWidth=1.4;fontSize=13;align=center;verticalAlign=middle;spacing=8;" vertex="1" parent="1">
          <mxGeometry x="380" y="435" width="220" height="200" as="geometry"/>
        </mxCell>
        <mxCell id="reconf_box" value="&lt;b&gt;Reconfiguration and Transformation&lt;/b&gt;&lt;hr&gt;Door sections rotation forward / backward&lt;hr&gt;Angular position change&lt;br&gt;Temporary stop due to pressure" style="rounded=0;whiteSpace=wrap;html=1;fillColor=#ffffff;strokeColor=#000000;strokeWidth=1.4;fontSize=13;align=center;verticalAlign=middle;spacing=8;" vertex="1" parent="1">
          <mxGeometry x="625" y="435" width="285" height="200" as="geometry"/>
        </mxCell>
        <mxCell id="act_box" value="&lt;b&gt;Actuators (motor)&lt;/b&gt;&lt;hr&gt;&lt;b&gt;IN&lt;/b&gt; motorOn, rotationForward &lt;b&gt;TO&lt;/b&gt; Plant" style="rounded=0;whiteSpace=wrap;html=1;fillColor=#ffffff;strokeColor=#000000;strokeWidth=1.4;fontSize=13;align=center;verticalAlign=middle;spacing=8;" vertex="1" parent="1">
          <mxGeometry x="935" y="435" width="165" height="200" as="geometry"/>
        </mxCell>
        <mxCell id="ind_box" value="&lt;b&gt;Indicators (traffic lights)&lt;/b&gt;&lt;hr&gt;&lt;b&gt;IN&lt;/b&gt; lightSideA, lightSideB &lt;b&gt;OUT&lt;/b&gt;" style="rounded=0;whiteSpace=wrap;html=1;fillColor=#ffffff;strokeColor=#000000;strokeWidth=1.4;fontSize=13;align=center;verticalAlign=middle;spacing=8;" vertex="1" parent="1">
          <mxGeometry x="1120" y="435" width="160" height="200" as="geometry"/>
        </mxCell>

        <mxCell id="controller_rect" value="" style="rounded=0;whiteSpace=wrap;html=1;fillColor=#ffffff;strokeColor=#000000;strokeWidth=1.6;" vertex="1" parent="1">
          <mxGeometry x="170" y="730" width="1080" height="255" as="geometry"/>
        </mxCell>
        <mxCell id="controller_label" value="&lt;b&gt;Controller&lt;/b&gt;" style="text;html=1;align=center;verticalAlign=middle;resizable=0;autosize=1;fontSize=18;fontStyle=1;" vertex="1" parent="1">
          <mxGeometry x="635" y="740" width="130" height="30" as="geometry"/>
        </mxCell>

        <mxCell id="adinput_box" value="&lt;b&gt;Analog/Digital Inputs&lt;/b&gt;&lt;hr&gt;&lt;b&gt;IN&lt;/b&gt; presenceSideA : BOOL&lt;br&gt;&lt;b&gt;IN&lt;/b&gt; presenceSideB : BOOL&lt;br&gt;&lt;b&gt;IN&lt;/b&gt; partitionPressure : BOOL&lt;br&gt;&lt;b&gt;IN&lt;/b&gt; rotateForwardButton : BOOL&lt;br&gt;&lt;b&gt;IN&lt;/b&gt; rotateBackwardButton : BOOL&lt;hr&gt;&lt;b&gt;IN&lt;/b&gt; doorAngle : REAL&lt;br&gt;&lt;b&gt;IN&lt;/b&gt; isRotating : BOOL" style="rounded=0;whiteSpace=wrap;html=1;fillColor=#ffffff;strokeColor=#000000;strokeWidth=1.4;fontSize=13;align=center;verticalAlign=middle;spacing=8;" vertex="1" parent="1">
          <mxGeometry x="195" y="780" width="350" height="185" as="geometry"/>
        </mxCell>
        <mxCell id="software_box" value="&lt;b&gt;Control Software (Info Transformation)&lt;/b&gt;&lt;hr&gt;MainControl&lt;hr&gt;StartRotationForward,&lt;br&gt;StartRotationBackward,&lt;br&gt;StopRotation, PauseRotation" style="rounded=0;whiteSpace=wrap;html=1;fillColor=#ffffff;strokeColor=#000000;strokeWidth=1.4;fontSize=13;align=center;verticalAlign=middle;spacing=8;" vertex="1" parent="1">
          <mxGeometry x="575" y="780" width="305" height="185" as="geometry"/>
        </mxCell>
        <mxCell id="adoutput_box" value="&lt;b&gt;Analog/Digital Outputs&lt;/b&gt;&lt;hr&gt;&lt;b&gt;OUT&lt;/b&gt; motorOn : BOOL&lt;br&gt;&lt;b&gt;OUT&lt;/b&gt; rotationForward : BOOL&lt;hr&gt;&lt;b&gt;OUT&lt;/b&gt; lightSideA : BOOL&lt;br&gt;&lt;b&gt;OUT&lt;/b&gt; lightSideB : BOOL" style="rounded=0;whiteSpace=wrap;html=1;fillColor=#ffffff;strokeColor=#000000;strokeWidth=1.4;fontSize=13;align=center;verticalAlign=middle;spacing=8;" vertex="1" parent="1">
          <mxGeometry x="915" y="780" width="310" height="185" as="geometry"/>
        </mxCell>

        <mxCell id="arr_users_actions" style="edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#000000;strokeWidth=1.5;endArrow=block;endFill=1;" edge="1" source="users" target="actions_box" parent="1">
          <mxGeometry relative="1" as="geometry"/>
        </mxCell>
        <mxCell id="arr_effects_sensation" style="edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#000000;strokeWidth=1.5;endArrow=block;endFill=1;" edge="1" source="effects_box" target="sensation_box" parent="1">
          <mxGeometry relative="1" as="geometry"/>
        </mxCell>
        <mxCell id="arr_sensation_users" style="edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#000000;strokeWidth=1.5;endArrow=block;endFill=1;" edge="1" source="sensation_box" target="users" parent="1">
          <mxGeometry relative="1" as="geometry">
            <Array as="points">
              <mxPoint x="1225" y="90"/>
              <mxPoint x="125" y="90"/>
            </Array>
          </mxGeometry>
        </mxCell>

        <mxCell id="arr_actions_controls" style="edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#000000;strokeWidth=1.5;endArrow=block;endFill=1;" edge="1" source="actions_box" target="ctrl_box" parent="1">
          <mxGeometry relative="1" as="geometry"/>
        </mxCell>
        <mxCell id="arr_effects_reconf" style="edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#000000;strokeWidth=1.5;startArrow=block;startFill=1;endArrow=block;endFill=1;" edge="1" source="effects_box" target="reconf_box" parent="1">
          <mxGeometry relative="1" as="geometry"/>
        </mxCell>
        <mxCell id="arr_indicators_sensation" style="edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#000000;strokeWidth=1.5;endArrow=block;endFill=1;" edge="1" source="ind_box" target="sensation_box" parent="1">
          <mxGeometry relative="1" as="geometry"/>
        </mxCell>
        <mxCell id="arr_reconf_sensors" style="edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#000000;strokeWidth=1.5;endArrow=block;endFill=1;" edge="1" source="reconf_box" target="sens_box" parent="1">
          <mxGeometry relative="1" as="geometry"/>
        </mxCell>
        <mxCell id="arr_actuators_reconf" style="edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#000000;strokeWidth=1.5;endArrow=block;endFill=1;" edge="1" source="act_box" target="reconf_box" parent="1">
          <mxGeometry relative="1" as="geometry"/>
        </mxCell>
        <mxCell id="arr_controls_inputs" style="edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#000000;strokeWidth=1.5;endArrow=block;endFill=1;" edge="1" source="ctrl_box" target="adinput_box" parent="1">
          <mxGeometry relative="1" as="geometry"/>
        </mxCell>
        <mxCell id="arr_sensors_inputs" style="edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#000000;strokeWidth=1.5;endArrow=block;endFill=1;" edge="1" source="sens_box" target="adinput_box" parent="1">
          <mxGeometry relative="1" as="geometry"/>
        </mxCell>
        <mxCell id="arr_inputs_software" style="edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#000000;strokeWidth=1.5;endArrow=block;endFill=1;" edge="1" source="adinput_box" target="software_box" parent="1">
          <mxGeometry relative="1" as="geometry"/>
        </mxCell>
        <mxCell id="arr_software_outputs" style="edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#000000;strokeWidth=1.5;endArrow=block;endFill=1;" edge="1" source="software_box" target="adoutput_box" parent="1">
          <mxGeometry relative="1" as="geometry"/>
        </mxCell>
        <mxCell id="arr_outputs_actuators" style="edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#000000;strokeWidth=1.5;endArrow=block;endFill=1;" edge="1" source="adoutput_box" target="act_box" parent="1">
          <mxGeometry relative="1" as="geometry"/>
        </mxCell>
        <mxCell id="arr_outputs_indicators" style="edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#000000;strokeWidth=1.5;endArrow=block;endFill=1;" edge="1" source="adoutput_box" target="ind_box" parent="1">
          <mxGeometry relative="1" as="geometry"/>
        </mxCell>
      </root>
    </mxGraphModel>
  </diagram>
</mxfile>
"""
    DRAWIO_PATH.write_text(drawio_xml, encoding="utf-8")


def wrap_text(text: str, limit: int) -> list[str]:
    words = text.split()
    lines: list[str] = []
    current = ""
    for word in words:
        candidate = word if not current else f"{current} {word}"
        if len(candidate) <= limit:
            current = candidate
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines or [text]


def draw_multiline(
    draw: ImageDraw.ImageDraw,
    xy: tuple[int, int],
    text: str,
    font: ImageFont.ImageFont,
    fill: str,
    line_gap: int = 4,
) -> None:
    x, y = xy
    for line in text.split("\n"):
        draw.text((x, y), line, font=font, fill=fill)
        bbox = draw.textbbox((x, y), line, font=font)
        y = bbox[3] + line_gap


def render_cpd_png() -> None:
    image = Image.new("RGB", (1400, 1000), "#ffffff")
    draw = ImageDraw.Draw(image)

    font_title = load_font(24, bold=True)
    font_label = load_font(20, bold=True)
    font_text = load_font(16, bold=False)
    font_small = load_font(14, bold=False)

    draw.rounded_rectangle((100, 20, 1300, 300), radius=60, outline="#666666", width=3, fill="#f5f5f5")
    draw.text((620, 35), "Environment", font=font_title, fill="#111111")

    draw.ellipse((106, 85, 144, 123), outline="#444444", width=2)
    draw.line((125, 123, 125, 165), fill="#444444", width=2)
    draw.line((105, 137, 145, 137), fill="#444444", width=2)
    draw.line((125, 165, 105, 195), fill="#444444", width=2)
    draw.line((125, 165, 145, 195), fill="#444444", width=2)
    draw.text((92, 202), "Users", font=font_small, fill="#111111")

    boxes = {
        "actions": (200, 90, 620, 280),
        "effects": (650, 120, 880, 260),
        "sensation": (920, 110, 1270, 280),
        "controls": (110, 380, 355, 590),
        "sensors": (380, 380, 600, 590),
        "reconf": (625, 380, 910, 590),
        "actuators": (935, 380, 1110, 590),
        "indicators": (1130, 380, 1280, 590),
        "inputs": (195, 680, 545, 940),
        "software": (575, 680, 885, 940),
        "outputs": (915, 680, 1235, 940),
    }

    for left, top, right, bottom in boxes.values():
        draw.rounded_rectangle((left, top, right, bottom), radius=18, outline="#333333", width=3, fill="#f5f5f5")

    draw.rounded_rectangle((90, 340, 1300, 620), radius=12, outline="#333333", width=3, fill="#ffffff")
    draw.text((640, 350), "Plant", font=font_title, fill="#111111")

    draw.rounded_rectangle((170, 650, 1250, 970), radius=12, outline="#333333", width=3, fill="#ffffff")
    draw.text((620, 660), "Controller", font=font_title, fill="#111111")

    draw.text((330, 112), "Actions:", font=font_label, fill="#111111")
    draw_multiline(
        draw,
        (228, 150),
        "ApproachSideA, ApproachSideB\nLeaveSideA, LeaveSideB\nPressPartition, ReleasePartition\nPressForwardButton, ReleaseForwardButton\nPressBackwardButton, ReleaseBackwardButton",
        font_text,
        "#111111",
    )

    draw.text((700, 135), "Effects", font=font_label, fill="#111111")
    draw_multiline(
        draw,
        (675, 175),
        "Passage availability\nVisible door movement\nin selected direction\nTemporary stop under pressure",
        font_small,
        "#111111",
    )

    draw.text((1010, 122), "Sensation:", font=font_label, fill="#111111")
    draw_multiline(
        draw,
        (955, 162),
        "User sees traffic lights\nand perceives forward /\nbackward rotation or\ntemporary stop",
        font_text,
        "#111111",
    )

    draw.text((136, 398), "Controls:", font=font_label, fill="#111111")
    draw_multiline(
        draw,
        (136, 435),
        "presenceSideA, presenceSideB OUT\npartitionPressure OUT\n----------------------------\nrotateForwardButton OUT\nrotateBackwardButton OUT",
        font_text,
        "#111111",
    )

    draw.text((395, 398), "Sensors:", font=font_label, fill="#111111")
    draw_multiline(draw, (392, 438), "doorAngle\nisRotating", font_text, "#111111")

    draw.text((596, 398), "Reconfiguration\nand Transformation:", font=font_label, fill="#111111")
    draw_multiline(
        draw,
        (645, 458),
        "Door rotation forward/backward\nAngular position change\nTemporary stop due to pressure",
        font_text,
        "#111111",
    )

    draw.text((850, 398), "Actuators:", font=font_label, fill="#111111")
    draw_multiline(draw, (952, 445), "motor\nmotorOn\nrotationForward", font_text, "#111111")

    draw.text((1084, 398), "Indicators:", font=font_label, fill="#111111")
    draw_multiline(draw, (1068, 438), "traffic lights\nlightSideA\nlightSideB", font_text, "#111111")

    draw.text((252, 708), "Analog/Digital Inputs:", font=font_label, fill="#111111")
    draw_multiline(
        draw,
        (225, 748),
        "presenceSideA, presenceSideB\npartitionPressure\nrotateForwardButton\nrotateBackwardButton\n----------------------------\ndoorAngle, isRotating",
        font_small,
        "#111111",
    )

    draw.text((560, 708), "Control Software\n(Info Transformation):", font=font_label, fill="#111111")
    draw_multiline(
        draw,
        (602, 780),
        "control processes\n(MainControl,\nStartRotationForward,\nStartRotationBackward,\nStopRotation,\nPauseRotation)",
        font_small,
        "#111111",
    )

    draw.text((874, 708), "Analog/Digital Outputs:", font=font_label, fill="#111111")
    draw_multiline(
        draw,
        (948, 748),
        "motorOn\nrotationForward\n----------------------------\nlightSideA\nlightSideB",
        font_text,
        "#111111",
    )

    def arrow(points: list[tuple[int, int]], width: int = 3) -> None:
        draw.line(points, fill="#333333", width=width)
        x1, y1 = points[-2]
        x2, y2 = points[-1]
        dx = x2 - x1
        dy = y2 - y1
        if dx == dy == 0:
            return
        if abs(dx) > abs(dy):
            sign = 1 if dx > 0 else -1
            draw.polygon([(x2, y2), (x2 - 14 * sign, y2 - 7), (x2 - 14 * sign, y2 + 7)], fill="#333333")
        else:
            sign = 1 if dy > 0 else -1
            draw.polygon([(x2, y2), (x2 - 7, y2 - 14 * sign), (x2 + 7, y2 - 14 * sign)], fill="#333333")

    arrow([(145, 137), (170, 137), (170, 185), (200, 185)])
    arrow([(880, 190), (920, 190)])
    arrow([(980, 100), (980, 70), (125, 70), (125, 85)])
    arrow([(355, 485), (245, 485), (245, 380)])
    arrow([(765, 260), (765, 380)])
    arrow([(1270, 380), (1270, 320), (1130, 320), (1130, 280)])
    arrow([(625, 485), (600, 485)])
    arrow([(935, 485), (910, 485)])
    arrow([(490, 590), (490, 680)])
    arrow([(245, 590), (330, 590), (330, 680)])
    arrow([(545, 810), (575, 810)])
    arrow([(885, 810), (915, 810)])
    arrow([(1000, 680), (1000, 590)])
    arrow([(1160, 680), (1160, 590)])

    image.save(PNG_PATH)


def paragraph_text(elem: etree._Element) -> str:
    return "".join(elem.xpath(".//w:t/text()", namespaces=NS)).strip()


def build_run_like(template_run: etree._Element | None, text: str, bold: bool | None = None) -> etree._Element:
    if template_run is None:
        run = etree.Element(w_tag("r"))
        rpr = etree.SubElement(run, w_tag("rPr"))
        size = etree.SubElement(rpr, w_tag("sz"))
        size.set(w_tag("val"), "28")
        size_cs = etree.SubElement(rpr, w_tag("szCs"))
        size_cs.set(w_tag("val"), "28")
    else:
        run = copy.deepcopy(template_run)
        for child in list(run):
            if child.tag == w_tag("t"):
                run.remove(child)
        rpr = run.find(w_tag("rPr"))
        if rpr is None:
            rpr = etree.SubElement(run, w_tag("rPr"))

    if bold is not None:
        bold_nodes = rpr.findall(w_tag("b"))
        if bold and not bold_nodes:
            etree.SubElement(rpr, w_tag("b"))
        if not bold:
            for node in bold_nodes:
                rpr.remove(node)

    text_node = etree.SubElement(run, w_tag("t"))
    if text.startswith(" ") or text.endswith(" "):
        text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text_node.text = text
    return run


def set_paragraph_runs(
    paragraph: etree._Element,
    run_specs: list[tuple[str, bool | None]],
    source_run: etree._Element | None = None,
) -> None:
    for child in list(paragraph):
        if child.tag == w_tag("r"):
            paragraph.remove(child)
    template_run = source_run if source_run is not None else paragraph.find(w_tag("r"))
    if template_run is not None and template_run.getparent() is paragraph:
        paragraph.remove(template_run)
    for text, bold in run_specs:
        paragraph.append(build_run_like(template_run, text, bold))


def make_paragraph_like(template: etree._Element, run_specs: list[tuple[str, bool | None]]) -> etree._Element:
    new_p = copy.deepcopy(template)
    source_run = new_p.find(w_tag("r"))
    set_paragraph_runs(new_p, run_specs, source_run=source_run)
    return new_p


def set_cell_text(cell: etree._Element, run_specs: list[tuple[str, bool | None]]) -> None:
    paragraph = cell.find(f".//{w_tag('p')}")
    if paragraph is None:
        paragraph = etree.SubElement(cell, w_tag("p"))
    template_run = paragraph.find(w_tag("r"))
    set_paragraph_runs(paragraph, run_specs, source_run=template_run)


def fix_docx_text() -> None:
    with zipfile.ZipFile(DOCX_PATH, "r") as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    doc_root = etree.fromstring(files["word/document.xml"])
    body = doc_root.find(w_tag("body"))
    if body is None:
        raise RuntimeError("word/document.xml has no body")

    def find_paragraph_contains(text: str) -> etree._Element:
        for item in body:
            if item.tag == w_tag("p") and text in paragraph_text(item):
                return item
        raise RuntimeError(f"paragraph not found: {text}")

    def maybe_find_paragraph_contains(text: str) -> etree._Element | None:
        for item in body:
            if item.tag == w_tag("p") and text in paragraph_text(item):
                return item
        return None

    def replace_paragraph_text(text: str, new_text: str) -> etree._Element:
        paragraph = maybe_find_paragraph_contains(text)
        if paragraph is None:
            paragraph = maybe_find_paragraph_contains(new_text)
        if paragraph is None:
            raise RuntimeError(f"paragraph not found for replace: {text}")
        set_paragraph_runs(paragraph, [(new_text, None)], source_run=paragraph.find(w_tag("r")))
        return paragraph

    def insert_after(anchor: etree._Element, new_text: str) -> etree._Element:
        paragraph = make_paragraph_like(anchor, [(new_text, None)])
        anchor.addnext(paragraph)
        return paragraph

    def find_all_exact(text: str) -> list[etree._Element]:
        matches: list[etree._Element] = []
        for item in body:
            if item.tag == w_tag("p") and paragraph_text(item) == text:
                matches.append(item)
        return matches

    term_para = find_paragraph_contains("секция – Section (A, B, C)")
    normal_template = find_paragraph_contains("Вращающаяся дверь — это трёхсекционная")
    body.replace(term_para, make_paragraph_like(normal_template, [("секция – Section (A, B, C)", None)]))

    approach_para = find_paragraph_contains("Оператор может имитировать подход пользователя")
    if maybe_find_paragraph_contains("LeaveSideA / LeaveSideB") is None:
        leave_para = make_paragraph_like(
            approach_para,
            [("Оператор может имитировать уход пользователя с каждой стороны (кнопки LeaveSideA / LeaveSideB).", None)],
        )
        approach_para.addnext(leave_para)

    controller_para = find_paragraph_contains("Светофоры переключаются в зависимости от состояния двери")
    if maybe_find_paragraph_contains("датчик isRotating должен принимать значение TRUE") is None:
        sensor_logic_para = make_paragraph_like(
            controller_para,
            [("При включённом двигателе датчик isRotating должен принимать значение TRUE, а doorAngle должен изменяться с постоянной скоростью.", None)],
        )
        controller_para.addnext(sensor_logic_para)

    merged_para = maybe_find_paragraph_contains("Установить индикаторы вручную.Controls:")
    if merged_para is not None:
        set_paragraph_runs(merged_para, [("Установить индикаторы вручную.", None)], source_run=merged_para.find(w_tag("r")))
        controls_label = make_paragraph_like(merged_para, [("Controls:", None)])
        merged_para.addnext(controls_label)

    elsif_para = maybe_find_paragraph_contains("ELSIF")
    if elsif_para is not None:
        set_paragraph_runs(elsif_para, [("      ELSE", None)], source_run=elsif_para.find(w_tag("r")))

    table = body.find(f".//{w_tag('tbl')}")
    if table is None:
        raise RuntimeError("requirements table not found")
    rows = table.findall(w_tag("tr"))
    if len(rows) < 3:
        raise RuntimeError("unexpected requirements table size")

    existing_rows = []
    blank_row = None
    for row in rows[1:]:
        first_cell = row.find(w_tag("tc"))
        if first_cell is not None:
            first_text = "".join(first_cell.xpath(".//w:t/text()", namespaces=NS)).strip()
            existing_rows.append((row, first_text))
            if not first_text:
                blank_row = row
    for row, text in existing_rows:
        if text in {
            "При включённом двигателе датчик isRotating должен быть активен",
            "При вращении значение doorAngle должно изменяться",
        }:
            table.remove(row)

    rows = table.findall(w_tag("tr"))
    template_row = copy.deepcopy(rows[2])
    if blank_row is not None and blank_row.getparent() is table:
        table.remove(blank_row)

    present_descriptions = set()
    for row in table.findall(w_tag("tr"))[1:]:
        first_cell = row.find(w_tag("tc"))
        if first_cell is not None:
            present_descriptions.add("".join(first_cell.xpath(".//w:t/text()", namespaces=NS)).strip())

    ensured_rows = [
        [
            [("Индикаторы должны гореть зелёным в исходном состоянии (дверь стоит, вход разрешён)", None)],
            [("NOT motorOn AND NOT partitionPressure", None)],
            [("FALSE", True)],
            [("TRUE", True)],
            [("TRUE", True)],
            [("lightSideA AND lightSideB", None)],
            [("TRUE", True)],
        ],
        [
            [("При включённом двигателе датчик isRotating должен быть активен", None)],
            [("motorOn", None)],
            [("FALSE", True)],
            [("TRUE", True)],
            [("TRUE", True)],
            [("isRotating", None)],
            [("TRUE", True)],
        ],
        [
            [("При вращении значение doorAngle должно изменяться", None)],
            [("motorOn", None)],
            [("FALSE", True)],
            [("TRUE", True)],
            [("TRUE", True)],
            [("doorAngle >= 0.0", None)],
            [("TRUE", True)],
        ],
    ]

    for specs in ensured_rows:
        description = specs[0][0][0]
        if description in present_descriptions:
            continue
        row = copy.deepcopy(template_row)
        cells = row.findall(w_tag("tc"))
        for cell, spec in zip(cells, specs):
            set_cell_text(cell, spec)
        table.append(row)
        present_descriptions.add(description)

    replace_paragraph_text(
        "Вращающаяся дверь — это трёхсекционная",
        "Вращающаяся дверь — это трёхсекционная (SECTION_COUNT = 3) дверь, установленная в проёме здания, обеспечивающая проход людей с двух сторон (SideA / SideB). В нормальном (исходном) состоянии дверь неподвижна. При приближении пользователя к одной из сторон (датчики presenceSideA / presenceSideB) двигатель (motorOn) запускает вращение двери с заданной скоростью (ROTATION_SPEED). Вращение продолжается в течение заданного времени (ROTATION_TIME). Если за время вращения повторно срабатывает один из сенсоров, таймер сбрасывается и отсчёт начинается заново. Вращение останавливается по истечении таймера. При оказании давления на секционные перегородки (датчик partitionPressure) вращение приостанавливается на время PAUSE_DURATION, после чего возобновляется, если таймер ещё не истёк. Для ручного управления предусмотрены специальные кнопки rotateForwardButton / rotateBackwardButton, позволяющие вращать дверь вперёд или назад, пока соответствующая кнопка удерживается нажатой.",
    )

    press_para = find_paragraph_contains("PressPartition / ReleasePartition")
    if maybe_find_paragraph_contains("rotateForwardButton / rotateBackwardButton для ручного вращения") is None:
        insert_after(
            press_para,
            "Оператор может нажимать специальные кнопки rotateForwardButton / rotateBackwardButton для ручного вращения двери вперёд или назад.",
        )

    replace_paragraph_text(
        "Двигатель запускается при срабатывании одного из двух сенсоров движения.",
        "Двигатель запускается при срабатывании одного из двух сенсоров движения или при нажатии специальных кнопок rotateForwardButton / rotateBackwardButton.",
    )

    if maybe_find_paragraph_contains("дверь должна вращаться вперёд") is None:
        direction_para = insert_after(
            controller_para,
            "При нажатии rotateForwardButton дверь должна вращаться вперёд, а при нажатии rotateBackwardButton — назад.",
        )
        insert_after(
            direction_para,
            "Одновременное нажатие rotateForwardButton и rotateBackwardButton недопустимо и не должно запускать двигатель.",
        )

    replace_paragraph_text(
        "Environment: ApproachSideA, ApproachSideB, LeaveSideA, LeaveSideB, PressPartition, ReleasePartition",
        "Environment: ApproachSideA, ApproachSideB, LeaveSideA, LeaveSideB, PressPartition, ReleasePartition, PressForwardButton, ReleaseForwardButton, PressBackwardButton, ReleaseBackwardButton",
    )

    for paragraph in find_all_exact("Controls: presenceSideA, presenceSideB, partitionPressure"):
        set_paragraph_runs(
            paragraph,
            [("Controls: presenceSideA, presenceSideB, partitionPressure, rotateForwardButton, rotateBackwardButton", None)],
            source_run=paragraph.find(w_tag("r")),
        )

    for paragraph in find_all_exact("Actuators: motorOn"):
        set_paragraph_runs(
            paragraph,
            [("Actuators: motorOn, rotationForward", None)],
            source_run=paragraph.find(w_tag("r")),
        )

    todo_para = maybe_find_paragraph_contains("TODO: Add ability to rotate door")
    if todo_para is not None:
        set_paragraph_runs(
            todo_para,
            [("Ручное вращение поддерживается кнопками rotateForwardButton / rotateBackwardButton.", None)],
            source_run=todo_para.find(w_tag("r")),
        )

    manual_para = maybe_find_paragraph_contains("Включить/выключить двигатель вручную")
    if manual_para is not None:
        set_paragraph_runs(
            manual_para,
            [("Нажать/отпустить rotateForwardButton для вращения двери вперёд.", None)],
            source_run=manual_para.find(w_tag("r")),
        )
        backward_para = manual_para.getnext()
        if backward_para is None or paragraph_text(backward_para) != "Нажать/отпустить rotateBackwardButton для вращения двери назад.":
            backward_para = insert_after(manual_para, "Нажать/отпустить rotateBackwardButton для вращения двери назад.")
            pressure_manual_para = insert_after(
                backward_para,
                "При активном давлении на перегородку нажатие любой из кнопок ручного вращения должно приводить к ошибке.",
            )
            insert_after(
                pressure_manual_para,
                "Одновременное нажатие rotateForwardButton и rotateBackwardButton недопустимо.",
            )

    repeated_controls_para = maybe_find_paragraph_contains("При срабатывании сенсора дверь вращается ROTATION_TIME.")
    if repeated_controls_para is not None and maybe_find_paragraph_contains("Нажатие rotateForwardButton / rotateBackwardButton запускает ручное вращение") is None:
        insert_after(
            repeated_controls_para,
            "Нажатие rotateForwardButton / rotateBackwardButton запускает ручное вращение в выбранном направлении, пока кнопка удерживается нажатой.",
        )

    classified_para = find_paragraph_contains("CLASSIFIED")
    code_template = classified_para.getnext()
    requirements_para = find_paragraph_contains("Требования")
    if code_template is None:
        raise RuntimeError("code block start not found after CLASSIFIED")

    next_after_code = requirements_para
    to_remove: list[etree._Element] = []
    removing = False
    for item in body:
        if item is code_template:
            removing = True
        if item is requirements_para:
            break
        if removing:
            to_remove.append(item)
    for item in to_remove:
        body.remove(item)

    updated_code_lines = [
        ("VAR_GLOBAL (* VAR_INPUT *)", True),
        ("\t(* Presence sensors *)", True),
        ("         presenceSideA  : BOOL;", False),
        ("         presenceSideB  : BOOL;", False),
        ("\t(* Partition pressure sensor *)", True),
        ("         partitionPressure  : BOOL;", False),
        ("\t(* Manual rotation buttons *)", True),
        ("         rotateForwardButton  : BOOL;", False),
        ("         rotateBackwardButton  : BOOL;", False),
        ("\t(* Door angle sensor *)", True),
        ("         doorAngle  : REAL;", False),
        ("\t(* Rotation status *)", True),
        ("         isRotating  : BOOL;", False),
        ("(* END_VAR", True),
        ("VAR *) (* VAR_OUTPUT *)", True),
        ("\t(* Motor control *)", True),
        ("         motorOn  : BOOL;", False),
        ("         rotationForward  : BOOL;", False),
        ("\t(* Traffic lights *)", True),
        ("         lightSideA  : BOOL;", False),
        ("         lightSideB  : BOOL;", False),
        ("END_VAR", True),
        ("", False),
        ("PROGRAM Plant (* RevolvingDoor *)", True),
        ("(* Plant *)", True),
        ("PROCESS Init", True),
        ("STATE begin", False),
        ("(* inputs: *)", True),
        ("presenceSideA := FALSE;", False),
        ("presenceSideB := FALSE;", False),
        ("partitionPressure := FALSE;", False),
        ("rotateForwardButton := FALSE;", False),
        ("rotateBackwardButton := FALSE;", False),
        ("doorAngle := 0.0;", False),
        ("isRotating := FALSE;", False),
        ("(* outputs: *)", True),
        ("motorOn := FALSE;", False),
        ("rotationForward := TRUE;", False),
        ("lightSideA := TRUE;", False),
        ("lightSideB := TRUE;", False),
        ("START PROCESS MotorSim;", False),
        ("START PROCESS DoorRotationSim;", False),
        ("START PROCESS PressureSim;", False),
        ("STOP;", False),
        ("END_STATE", False),
        ("END_PROCESS", False),
        ("", False),
        ("PROCESS MotorSim", True),
        ("VAR CONSTANT", True),
        ("ROTATION_SPEED : REAL := 2.0;", False),
        ("END_VAR", True),
        ("STATE check_motor LOOPED", False),
        ("IF motorOn THEN", False),
        ("isRotating := TRUE;", False),
        ("ELSE", False),
        ("isRotating := FALSE;", False),
        ("END_IF", False),
        ("END_STATE", False),
        ("END_PROCESS", False),
        ("", False),
        ("PROCESS DoorRotationSim", True),
        ("VAR CONSTANT", True),
        ("ROTATION_SPEED : REAL := 2.0;", False),
        ("MAX_ANGLE : REAL := 360.0;", False),
        ("END_VAR", True),
        ("STATE check_rotation LOOPED", False),
        ("IF isRotating THEN", False),
        ("IF rotationForward THEN", False),
        ("doorAngle := doorAngle + ROTATION_SPEED;", False),
        ("ELSE", False),
        ("doorAngle := doorAngle - ROTATION_SPEED;", False),
        ("END_IF", False),
        ("END_IF", False),
        ("IF doorAngle >= MAX_ANGLE THEN", False),
        ("doorAngle := doorAngle - MAX_ANGLE;", False),
        ("END_IF", False),
        ("IF doorAngle < 0.0 THEN", False),
        ("doorAngle := doorAngle + MAX_ANGLE;", False),
        ("END_IF", False),
        ("END_STATE", False),
        ("END_PROCESS", False),
        ("", False),
        ("PROCESS PressureSim", True),
        ("VAR CONSTANT", True),
        ("PAUSE_DURATION : REAL := 30.0;", False),
        ("END_VAR", True),
        ("VAR", True),
        ("pauseCounter : REAL := 0.0;", False),
        ("END_VAR", True),
        ("STATE check_pressure LOOPED", False),
        ("IF partitionPressure THEN", False),
        ("pauseCounter := PAUSE_DURATION;", False),
        ("END_IF", False),
        ("IF pauseCounter > 0.0 THEN", False),
        ("pauseCounter := pauseCounter - 1.0;", False),
        ("END_IF", False),
        ("END_STATE", False),
        ("END_PROCESS", False),
        ("END_PROGRAM", True),
        ("", False),
        ("PROGRAM Controller", True),
        ("VAR (* VAR_INPUT *)", True),
        ("(* Presence sensors *)", True),
        ("presenceSideA : BOOL;", False),
        ("presenceSideB : BOOL;", False),
        ("(* Partition pressure sensor *)", True),
        ("partitionPressure : BOOL;", False),
        ("(* Manual rotation buttons *)", True),
        ("rotateForwardButton : BOOL;", False),
        ("rotateBackwardButton : BOOL;", False),
        ("(* Door angle sensor *)", True),
        ("doorAngle : REAL;", False),
        ("(* Rotation status *)", True),
        ("isRotating : BOOL;", False),
        ("END_VAR", True),
        ("VAR (* VAR_OUTPUT *)", True),
        ("(* Motor control *)", True),
        ("motorOn : BOOL;", False),
        ("rotationForward : BOOL;", False),
        ("(* Traffic lights *)", True),
        ("lightSideA : BOOL;", False),
        ("lightSideB : BOOL;", False),
        ("END_VAR", True),
        ("PROCESS MainControl", True),
        ("VAR CONSTANT", True),
        ("ROTATION_TIME : INT := 50;", False),
        ("END_VAR", True),
        ("VAR", True),
        ("timer : INT := 0;", False),
        ("manualActive : BOOL := FALSE;", False),
        ("manualDirectionForward : BOOL := TRUE;", False),
        ("END_VAR", True),
        ("STATE idle", False),
        ("IF rotateForwardButton AND NOT rotateBackwardButton THEN", False),
        ("manualActive := TRUE;", False),
        ("manualDirectionForward := TRUE;", False),
        ("START PROCESS StartRotationForward;", False),
        ("SET STATE manualRotation;", False),
        ("ELSIF rotateBackwardButton AND NOT rotateForwardButton THEN", False),
        ("manualActive := TRUE;", False),
        ("manualDirectionForward := FALSE;", False),
        ("START PROCESS StartRotationBackward;", False),
        ("SET STATE manualRotation;", False),
        ("ELSIF (presenceSideA OR presenceSideB) THEN", False),
        ("manualActive := FALSE;", False),
        ("manualDirectionForward := TRUE;", False),
        ("START PROCESS StartRotationForward;", False),
        ("timer := ROTATION_TIME;", False),
        ("SET STATE rotating;", False),
        ("END_IF", False),
        ("END_STATE", False),
        ("STATE rotating LOOPED", False),
        ("IF (presenceSideA OR presenceSideB) THEN", False),
        ("timer := ROTATION_TIME;", False),
        ("END_IF", False),
        ("IF partitionPressure THEN", False),
        ("START PROCESS PauseRotation;", False),
        ("SET STATE paused;", False),
        ("END_IF", False),
        ("timer := timer - 1;", False),
        ("IF timer <= 0 THEN", False),
        ("START PROCESS StopRotation;", False),
        ("SET STATE idle;", False),
        ("END_IF", False),
        ("END_STATE", False),
        ("STATE manualRotation LOOPED", False),
        ("IF partitionPressure THEN", False),
        ("START PROCESS PauseRotation;", False),
        ("SET STATE paused;", False),
        ("ELSIF rotateForwardButton AND rotateBackwardButton THEN", False),
        ("START PROCESS StopRotation;", False),
        ("manualActive := FALSE;", False),
        ("SET STATE idle;", False),
        ("ELSIF manualDirectionForward AND rotateForwardButton THEN", False),
        ("START PROCESS StartRotationForward;", False),
        ("ELSIF (NOT manualDirectionForward) AND rotateBackwardButton THEN", False),
        ("START PROCESS StartRotationBackward;", False),
        ("ELSE", False),
        ("START PROCESS StopRotation;", False),
        ("manualActive := FALSE;", False),
        ("SET STATE idle;", False),
        ("END_IF", False),
        ("END_STATE", False),
        ("STATE paused", False),
        ("IF (PROCESS PauseRotation IN STATE STOP) THEN", False),
        ("IF manualActive THEN", False),
        ("IF manualDirectionForward AND rotateForwardButton THEN", False),
        ("START PROCESS StartRotationForward;", False),
        ("SET STATE manualRotation;", False),
        ("ELSIF (NOT manualDirectionForward) AND rotateBackwardButton THEN", False),
        ("START PROCESS StartRotationBackward;", False),
        ("SET STATE manualRotation;", False),
        ("ELSE", False),
        ("START PROCESS StopRotation;", False),
        ("manualActive := FALSE;", False),
        ("SET STATE idle;", False),
        ("END_IF", False),
        ("ELSIF timer > 0 THEN", False),
        ("START PROCESS StartRotationForward;", False),
        ("SET STATE rotating;", False),
        ("ELSE", False),
        ("START PROCESS StopRotation;", False),
        ("SET STATE idle;", False),
        ("END_IF", False),
        ("END_IF", False),
        ("END_STATE", False),
        ("END_PROCESS", False),
        ("", False),
        ("PROCESS StartRotationForward", True),
        ("STATE init", False),
        ("motorOn := TRUE;", False),
        ("rotationForward := TRUE;", False),
        ("lightSideA := TRUE;", False),
        ("lightSideB := TRUE;", False),
        ("STOP;", False),
        ("END_STATE", False),
        ("END_PROCESS", False),
        ("", False),
        ("PROCESS StartRotationBackward", True),
        ("STATE init", False),
        ("motorOn := TRUE;", False),
        ("rotationForward := FALSE;", False),
        ("lightSideA := TRUE;", False),
        ("lightSideB := TRUE;", False),
        ("STOP;", False),
        ("END_STATE", False),
        ("END_PROCESS", False),
        ("", False),
        ("PROCESS StopRotation", True),
        ("STATE init", False),
        ("motorOn := FALSE;", False),
        ("rotationForward := TRUE;", False),
        ("lightSideA := TRUE;", False),
        ("lightSideB := TRUE;", False),
        ("STOP;", False),
        ("END_STATE", False),
        ("END_PROCESS", False),
        ("", False),
        ("PROCESS PauseRotation", True),
        ("VAR CONSTANT", True),
        ("PAUSE_DURATION : INT := 30;", False),
        ("END_VAR", True),
        ("VAR", True),
        ("counter : INT := 0;", False),
        ("END_VAR", True),
        ("STATE init", False),
        ("motorOn := FALSE;", False),
        ("lightSideA := FALSE;", False),
        ("lightSideB := FALSE;", False),
        ("counter := 0;", False),
        ("SET NEXT;", False),
        ("END_STATE", False),
        ("STATE waiting LOOPED", False),
        ("counter := counter + 1;", False),
        ("IF counter >= PAUSE_DURATION THEN", False),
        ("STOP;", False),
        ("END_IF", False),
        ("END_STATE", False),
        ("END_PROCESS", False),
        ("END_PROGRAM", True),
    ]

    inserted_before = next_after_code
    for text, bold in updated_code_lines:
        paragraph = make_paragraph_like(code_template, [(text, bold)])
        if inserted_before is None:
            body.append(paragraph)
        else:
            inserted_before.addprevious(paragraph)

    table_rows = table.findall(w_tag("tr"))
    header_row = table_rows[0]
    data_template = copy.deepcopy(table_rows[1])
    for row in table_rows[1:]:
        table.remove(row)

    rebuilt_rows = [
        [[("default", None)], [("TRUE", True)], [("FALSE", True)], [("TRUE", True)], [("TRUE", True)], [("TRUE", True)], [("TRUE", True)]],
        [[("Дверь должна начать вращаться при активации одного из двух сенсоров движения", None)], [("presenceSideA.RE OR presenceSideB.RE", None)], [("FALSE", True)], [("TRUE", True)], [("TRUE", True)], [("motorOn AND rotationForward", None)], [("TRUE", True)]],
        [[("Дверь должна вращаться в течение заданного времени ROTATION_TIME после последней активации сенсора", None)], [("presenceSideA.RE OR presenceSideB.RE", None)], [("FALSE", True)], [("tau(#ROTATION_TIME)", None)], [("TRUE", True)], [("motorOn AND rotationForward", None)], [("NOT motorOn", None)]],
        [[("При нажатии rotateForwardButton дверь должна начать вращаться вперёд", None)], [("rotateForwardButton.RE", None)], [("FALSE", True)], [("TRUE", True)], [("TRUE", True)], [("motorOn AND rotationForward", None)], [("TRUE", True)]],
        [[("При нажатии rotateBackwardButton дверь должна начать вращаться назад", None)], [("rotateBackwardButton.RE", None)], [("FALSE", True)], [("TRUE", True)], [("TRUE", True)], [("motorOn AND NOT rotationForward", None)], [("TRUE", True)]],
        [[("Одновременное нажатие rotateForwardButton и rotateBackwardButton не должно запускать двигатель", None)], [("rotateForwardButton AND rotateBackwardButton", None)], [("FALSE", True)], [("TRUE", True)], [("TRUE", True)], [("NOT motorOn", None)], [("TRUE", True)]],
        [[("При давлении на перегородку вращение должно быть приостановлено", None)], [("partitionPressure", None)], [("FALSE", True)], [("TRUE", True)], [("TRUE", True)], [("NOT motorOn", None)], [("TRUE", True)]],
        [[("После снятия давления вращение возобновляется через PAUSE_DURATION если таймер ещё не истёк", None)], [("partitionPressure.FE AND timer > 0", None)], [("FALSE", True)], [("tau(#PAUSE_DURATION)", None)], [("TRUE", True)], [("TRUE", True)], [("motorOn", None)]],
        [[("Индикатор на стороне A должен гореть зелёным при нормальном вращении", None)], [("motorOn AND NOT partitionPressure", None)], [("FALSE", True)], [("TRUE", True)], [("TRUE", True)], [("lightSideA", None)], [("TRUE", True)]],
        [[("Индикатор на стороне B должен гореть зелёным при нормальном вращении", None)], [("motorOn AND NOT partitionPressure", None)], [("FALSE", True)], [("TRUE", True)], [("TRUE", True)], [("lightSideB", None)], [("TRUE", True)]],
        [[("Индикаторы должны гореть красным при паузе (давление на перегородку)", None)], [("partitionPressure", None)], [("FALSE", True)], [("TRUE", True)], [("TRUE", True)], [("NOT lightSideA AND NOT lightSideB", None)], [("TRUE", True)]],
        [[("Дверь не должна вращаться при истёкшем таймере и отсутствии активации сенсоров", None)], [("TRUE", True)], [("FALSE", True)], [("TRUE", True)], [("TRUE", True)], [("NOT (motorOn AND timer <= 0)", None)], [("TRUE", True)]],
        [[("Индикаторы должны гореть зелёным в исходном состоянии (дверь стоит, вход разрешён)", None)], [("NOT motorOn AND NOT partitionPressure", None)], [("FALSE", True)], [("TRUE", True)], [("TRUE", True)], [("lightSideA AND lightSideB", None)], [("TRUE", True)]],
        [[("При включённом двигателе датчик isRotating должен быть активен", None)], [("motorOn", None)], [("FALSE", True)], [("TRUE", True)], [("TRUE", True)], [("isRotating", None)], [("TRUE", True)]],
        [[("При вращении значение doorAngle должно изменяться", None)], [("motorOn", None)], [("FALSE", True)], [("TRUE", True)], [("TRUE", True)], [("doorAngle >= 0.0", None)], [("TRUE", True)]],
    ]

    for specs in rebuilt_rows:
        row = copy.deepcopy(data_template)
        cells = row.findall(w_tag("tc"))
        for cell, spec in zip(cells, specs):
            set_cell_text(cell, spec)
        table.append(row)

    insert_cpd_image(body, files)

    files["word/document.xml"] = etree.tostring(doc_root, xml_declaration=True, encoding="UTF-8", standalone="yes")

    tmp_path = DOCX_PATH.with_suffix(".tmp.docx")
    with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    tmp_path.replace(DOCX_PATH)


def insert_cpd_image(body: etree._Element, files: dict[str, bytes]) -> None:
    if "word/media/RevolvingDoor_CyberPhysical.png" in files:
        files["word/media/RevolvingDoor_CyberPhysical.png"] = PNG_PATH.read_bytes()
        return

    sample_docx = ROOT / "Sample.docx"
    with zipfile.ZipFile(sample_docx, "r") as zin:
        sample_root = etree.fromstring(zin.read("word/document.xml"))

    image_para = sample_root.xpath("//w:p[.//wp:inline]", namespaces=NS)[0]
    image_para = copy.deepcopy(image_para)

    rels_root = etree.fromstring(files["word/_rels/document.xml.rels"])
    rel_ids = []
    for rel in rels_root.findall(f"{{{PKG_REL_NS}}}Relationship"):
        rel_id = rel.get("Id", "")
        if rel_id.startswith("rId") and rel_id[3:].isdigit():
            rel_ids.append(int(rel_id[3:]))
    next_rel_id = f"rId{max(rel_ids) + 1}"

    rel = etree.SubElement(
        rels_root,
        f"{{{PKG_REL_NS}}}Relationship",
        Id=next_rel_id,
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        Target="media/RevolvingDoor_CyberPhysical.png",
    )
    _ = rel

    for inline in image_para.xpath(".//wp:inline", namespaces=NS):
        inline.set("distB", "114300")
        inline.set("distT", "114300")
        inline.set("distL", "114300")
        inline.set("distR", "114300")

    width = "6217920"
    height = "4441371"

    for extent in image_para.xpath(".//wp:extent", namespaces=NS):
        extent.set("cx", width)
        extent.set("cy", height)
    for ext in image_para.xpath(".//a:ext", namespaces=NS):
        ext.set("cx", width)
        ext.set("cy", height)
    for doc_pr in image_para.xpath(".//wp:docPr", namespaces=NS):
        doc_pr.set("id", "1")
        doc_pr.set("name", "RevolvingDoor_CyberPhysical.png")
    for c_nv_pr in image_para.xpath(".//pic:cNvPr", namespaces=NS):
        c_nv_pr.set("id", "0")
        c_nv_pr.set("name", "RevolvingDoor_CyberPhysical.png")
    for blip in image_para.xpath(".//a:blip", namespaces=NS):
        blip.set(f"{{{R_NS}}}embed", next_rel_id)

    insert_after = None
    for item in body:
        if item.tag == w_tag("p") and "Indicators: lightSideA, lightSideB" in paragraph_text(item):
            insert_after = item
            break
    if insert_after is None:
        raise RuntimeError("indicator paragraph not found for image insertion")
    insert_after.addnext(image_para)

    files["word/_rels/document.xml.rels"] = etree.tostring(
        rels_root, xml_declaration=True, encoding="UTF-8", standalone="yes"
    )

    content_types = etree.fromstring(files["[Content_Types].xml"])
    defaults = content_types.findall(f"{{{CT_NS}}}Default")
    has_png = any(node.get("Extension") == "png" for node in defaults)
    if not has_png:
        etree.SubElement(
            content_types,
            f"{{{CT_NS}}}Default",
            Extension="png",
            ContentType="image/png",
        )
        files["[Content_Types].xml"] = etree.tostring(
            content_types, xml_declaration=True, encoding="UTF-8", standalone="yes"
        )

    files["word/media/RevolvingDoor_CyberPhysical.png"] = PNG_PATH.read_bytes()


def main() -> None:
    fix_drawio()
    render_cpd_png()
    fix_docx_text()


if __name__ == "__main__":
    main()
