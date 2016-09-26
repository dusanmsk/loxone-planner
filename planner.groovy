/**
 * @author dusan.zatkovsky , 9/22/16
 */

// https://mvnrepository.com/artifact/org.jopendocument/jOpenDocument
@Grapes(
    @Grab(group='org.jopendocument', module='jOpenDocument', version='1.3b1')
)


import org.jopendocument.dom.ODValueType
import org.jopendocument.dom.spreadsheet.MutableCell
import org.jopendocument.dom.spreadsheet.SpreadSheet

import java.awt.Color
import java.nio.file.Paths

class Main {

    def COLOR_OK = Color.GREEN
    def COLOR_ERROR = Color.RED
    def COLOR_WARN = Color.LIGHT_GRAY

    enum PLCType {
        Miniserver, Extension, Railduino
    }

    class Output {
        def id;
        def mandatory;
        def room;
        def place;
        def purpose;
        def note;
        def wattage;
        MutableCell cell
    }

    class Input {
        def id;
        def mandatory;
        def room;
        def place;
        def purpose;
        def note;
        MutableCell cell
    }


    class PLC {
        def id
        PLCType type
        int noOfInputs
        int noOfOutputs

        PLC(id, type, inputs, outputs) {
            this.id = id
            this.type = type
            this.noOfInputs = inputs
            this.noOfOutputs = outputs
        }
    }

    class MiniServer extends PLC {
        MiniServer(id) {
            super(id, PLCType.Miniserver, 8, 8)
        }
    }

    class Extension extends PLC {
        Extension(id) {
            super(id, PLCType.Extension, 12, 8)
        }
    }

    class Railduino extends PLC {
        Railduino(id) {
            super(id, PLCType.Railduino, 24, 12)
        }
        def groupped = [[3, 4, 5, 6], [9, 10, 11, 12]]
    }

    class Rack {
        MiniServer miniServer = new MiniServer('s');
        Extension extension1 = new Extension('e1');
        Railduino railduino1 = new Railduino('r1');
    }

    def toInt(String txt, defval) {
        if (txt.isNumber()) {
            return txt.toInteger()
        }
        return defval
    }


    def spreadSheet
    def outputs = [:]
    def inputs = [:]
    def connectedInputs = []
    def connectedOutputs = []

    def loadOutputs() {
        def outputSheet = spreadSheet.getSheet("outputs");
        (1..200).each { i ->
            def cell = outputSheet.getCellAt(0, i)
            def id = cell.getTextValue()
            if (id) {
                outputs[id] = new Output(
                        cell: cell,
                        id: id,
                        room: outputSheet.getCellAt(1, i).getTextValue(),
                        place: outputSheet.getCellAt(2, i).getTextValue(),
                        purpose: outputSheet.getCellAt(3, i).getTextValue(),
                        mandatory: !outputSheet.getCellAt(4, i).getTextValue().trim().isEmpty(),
                        note: outputSheet.getCellAt(5, i).getTextValue(),
                        wattage: toInt(outputSheet.getCellAt(6, i).getTextValue(), 0)
                )
            }
        }
    }

    def loadInputs() {
        def inputsSheet = spreadSheet.getSheet("inputs");
        (1..200).each { i ->
            def cell = inputsSheet.getCellAt(0, i)
            def id = cell.getTextValue()
            if (id) {
                inputs[id] = new Input(
                        cell: cell,
                        id: id,
                        room: inputsSheet.getCellAt(1, i).getTextValue(),
                        place: inputsSheet.getCellAt(2, i).getTextValue(),
                        purpose: inputsSheet.getCellAt(3, i).getTextValue(),
                        mandatory: !inputsSheet.getCellAt(4, i).getTextValue().trim().isEmpty(),
                        note: inputsSheet.getCellAt(5, i).getTextValue()
                )
            }
        }
    }

    def prepareColors() {
        // set every mandatory i/o red background - it will be greened if i/o is connected somewhere
        inputs.values().each { v ->
            v.cell.setBackgroundColor(COLOR_WARN)
            if (v.mandatory) {
                v.cell.setBackgroundColor(COLOR_ERROR)
            }
        }

        outputs.values().each { v ->
            v.cell.setBackgroundColor(COLOR_WARN)
            if (v.mandatory) {
                v.cell.setBackgroundColor(COLOR_ERROR)
            }
        }
    }

    def processMapping() {
        def mappingSheet = spreadSheet.getSheet("mapping");
        (1..200).each { i ->
            def mappingCell = mappingSheet.getCellAt(2, i)
            def tmp = mappingCell.getTextValue().trim()
            if (!tmp.isEmpty()) {
                def locations = []
                tmp.split(',').each { s ->
                    s = s.trim()
                    Input input = inputs[s]
                    if (input) {
                        if (input.id in connectedInputs) {
                            println "Error input already connected ${input.id}"
                            mappingCell.setBackgroundColor(COLOR_ERROR)
                        }
                        connectedInputs << input.id
                        input.cell.setBackgroundColor(COLOR_OK)
                        locations << "${input.room}:${input.place}:${input.purpose}"
                    } else {
                        println "Error input not found ${s}"
                        mappingCell.setBackgroundColor(COLOR_ERROR)
                    }
                }
                mappingSheet.getCellAt(1, i).setValue(locations.join(', '))
            }
        }

        (1..200).each { i ->
            def mappingCell = mappingSheet.getCellAt(6, i)
            def tmp = mappingCell.getTextValue().trim()
            if (!tmp.isEmpty()) {
                def locations = []
                tmp.split(',').each { s ->
                    s = s.trim()
                    Output output = outputs[s]
                    if (output) {
                        if (output.id in connectedOutputs) {
                            println "Error output already connected ${output.id}"
                            mappingCell.setBackgroundColor(COLOR_ERROR)
                        }
                        connectedOutputs << output.id
                        output.cell.setBackgroundColor(COLOR_OK)
                        float ampers = output.wattage / 230
                        mappingSheet.getCellAt(7, i).setValue(ampers, ODValueType.FLOAT, true, true)
                        locations << "${output.room}:${output.place}:${output.purpose}"
                    } else {
                        println "Error output not found ${s}"
                        mappingCell.setBackgroundColor(COLOR_ERROR)
                    }
                }
                mappingSheet.getCellAt(9, i).setValue(locations.join(', '))
            }
        }

    }

    def main(src, dst) {
        spreadSheet = SpreadSheet.createFromFile(src)
        loadInputs()
        loadOutputs()
        prepareColors()
        processMapping()
        spreadSheet.saveAs(dst)
    }
}

public static void main(String[] args) {
    def lastModified = null
    def src = Paths.get("${args[0]}").toFile()
    def dst = Paths.get("${args[1]}").toFile()
    while (true) {
        if (src.lastModified() != lastModified) {
            println "processing"
            lastModified = src.lastModified()

            def m = new Main()
            m.main(src, dst)
            println "saved\n\n"
        } else {
            Thread.sleep(5000)
        }
    }
}