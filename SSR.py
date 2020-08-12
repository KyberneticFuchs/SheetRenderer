bl_info = {
    "name": "GenText2",
    "blender": (2, 81, 0),
    "category": "Object",
}

import bpy
import time
import xlrd 

class ObjectMoveX(bpy.types.Operator):

    bl_idname = "object.gen_text"
    bl_label = "GenTex2"
    bl_options = {'REGISTER'} 
    
    SpreadSheet = "Path to the file"

    def execute(self, context): 
        
        scene = context.scene
        text = 'COFFE'
        
        bpy.data.scenes[0].render.filepath = SpreadSheet
        
        loc = ('/media/feetpaw/E6F47C3AF47C0ED5/UPWORK/Blender/font_things/Words_original.xlsx')
        wb = xlrd.open_workbook(loc)
        sheet = wb.sheet_by_index(0)  

        for i in range(sheet.nrows):
            if i == 165:
                return {'FINISHED'}
            bpy.data.scenes[0].render.filepath = bpy.data.scenes[0].render.filepath + sheet.cell_value(i, 1) + '/'
            bpy.context.scene.frame_current = 0
            #for k in range(0,13):
            layer0 = bpy.data.objects["Text.020"]
            layer1 = bpy.data.objects["Text.005"]
            layer2 = bpy.data.objects["Text.027"]
            layer3 = bpy.data.objects["Text.028"]
        
            layer0.data.body = sheet.cell_value(i, 1)
            layer1.data.body = sheet.cell_value(i, 1)
            layer2.data.body = sheet.cell_value(i, 1)
            layer3.data.body = sheet.cell_value(i, 1)

            bpy.ops.render.render(animation=True, use_viewport=True)
            bpy.data.scenes[0].render.filepath = '/media/feetpaw/E6F47C3AF47C0ED5/UPWORK/Blender/font_things/EXPORT/'
            
        return {'FINISHED'}
    
def register():
    bpy.utils.register_class(ObjectMoveX)


def unregister():
    bpy.utils.unregister_class(ObjectMoveX)

if __name__ == "__main__":
    register()
