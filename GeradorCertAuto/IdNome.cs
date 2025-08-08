using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GeradorCertAuto
{
    public class IdNome
    {
        public int Id { get; set; }
        public string Nome { get; set; }

        public IdNome()
        {}

        public IdNome(int id, string nome)
        {
            Id = id;
            Nome = nome;
        }

        public IdNome(ModeloCertificado modeloCertificado)
        {
            Id = modeloCertificado.Id ?? 0;
            Nome = modeloCertificado.Nome;
        }
    }
}
